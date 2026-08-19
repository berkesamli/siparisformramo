"use client";

import { useEffect, useMemo, useState } from "react";
import { FRAME_PROFILES, findProfile, boyLength, koliBoyText } from "@/data/catalog";
import { TECHNICAL_PRODUCTS, getTechnicalProduct } from "@/data/technical";
import { GLASS_TYPES, GLASS_SIZES, AYNA_SIZES, plateM2 } from "@/data/glass";
import { kurus, kesin, fmtQty, fmtPrice, fmtTL, sayi } from "@/lib/num";
import CustomerPicker from "@/components/CustomerPicker";
import TechnicalPicker from "@/components/TechnicalPicker";
import OrderTextImport, { type ParsedLine } from "@/components/OrderTextImport";

type Kind = "frame" | "glass" | "ayna" | "technical" | "other";

interface Row {
  id: number;
  kind: Kind;
  // frame
  code: string;
  unit: "metre" | "boy" | "koli";
  qty: string;
  usd: string;
  // Birim fiyat para birimi: varsayılan USD (liste fiyatı × kur). Müşteriyle
  // yuvarlak TL anlaşıldığında (47,85 → 47) ₺ seçilir, aynı kutuya TL yazılır.
  fx: "usd" | "tl";
  tl: string;
  // glass / ayna
  glassType: string;
  sizeIndex: number;
  plakaAdet: string;
  m2Price: string;
  // Müze camı fiyat para birimi: varsayılan € (liste × euro kuru).
  // ₺ seçilirse m² fiyatı doğrudan TL yazılır.
  glassFx: "eur" | "tl";
  // technical
  techCode: string;
  kartonKodu: string;
  kutuAdet: string;
  kutuPrice: string; // TL veya EUR (ürüne göre)
  // other
  name: string;
  otherQty: string;
  otherPrice: string;
  // Satır iskontosu (%) — her satırda ayrı oran olabilir (çerçevede %10,
  // teknik malzemede %5 gibi). Genel iskonto özetten ayrıca uygulanır.
  iskonto: string;
}

let rowSeq = 1;
const emptyRow = (): Row => ({
  id: rowSeq++,
  kind: "frame",
  code: "",
  unit: "metre",
  qty: "",
  usd: "",
  fx: "usd",
  tl: "",
  glassType: "duz",
  sizeIndex: 0,
  plakaAdet: "",
  m2Price: "",
  glassFx: "eur",
  techCode: "",
  kartonKodu: "",
  kutuAdet: "",
  kutuPrice: "",
  name: "",
  otherQty: "",
  otherPrice: "",
  iskonto: "",
});

// Miktar ve birim fiyatlar tam hassasiyetle (kesin) taşınır; yuvarlama
// yalnızca satır tutarında ve toplamlarda (kurus) yapılır. Bkz. lib/num.ts
const r2 = kurus;

// Teknik malzemeyi adıyla bulur ("10luk agraf" → "10'luk Agraf").
const sadeAd = (s: string) =>
  String(s || "")
    .toLowerCase()
    .replace(/[çÇ]/g, "c").replace(/[ğĞ]/g, "g").replace(/[ıİ]/g, "i")
    .replace(/[öÖ]/g, "o").replace(/[şŞ]/g, "s").replace(/[üÜ]/g, "u")
    .replace(/[^a-z0-9]/g, "");

function findTechnicalByName(q: string) {
  const k = sadeAd(q);
  if (k.length < 3) return undefined;
  return TECHNICAL_PRODUCTS.find((t) => {
    const n = sadeAd(t.name);
    return n === k || n.includes(k) || k.includes(n);
  });
}
const fmt = fmtTL;

interface ComputedLine {
  name: string;
  unitText: string;
  unitPriceTL: number;
  lineTotal: number;
}

/**
 * Satır iskontosunu uygular: indirim birim fiyata yansır ki fişte
 * miktar × birim fiyat = tutar her zaman birebir tutsun; oran ürün
 * adının yanına yazılır, müşteri de görebilir.
 */
function satirBitir(
  row: Row,
  name: string,
  unitText: string,
  unitPriceTL: number,
  qtyNum: number
): ComputedLine {
  const pct = Math.min(100, Math.max(0, sayi(row.iskonto) || 0));
  const birim = pct > 0 ? kesin(unitPriceTL * (1 - pct / 100)) : unitPriceTL;
  return {
    name: pct > 0 ? `${name} (%${fmtQty(pct)} isk.)` : name,
    unitText,
    unitPriceTL: birim,
    lineTotal: kurus(qtyNum * birim),
  };
}

function computeRow(row: Row, rate: number, euroRate: number): ComputedLine | null {
  if (row.kind === "frame") {
    const profile = findProfile(row.code);
    const qty = sayi(row.qty) || 0;
    if (!row.code.trim() || qty <= 0) return null;
    const usd = sayi(row.usd) || profile?.priceUSD || 0;
    // ₺ seçiliyse kutudaki TL fiyat geçerli; USD'de (veya TL boşsa) USD × kur
    const tlManuel = sayi(row.tl) || 0;
    const unitPriceTL =
      row.fx === "tl" && tlManuel > 0 ? kesin(tlManuel) : kesin(usd * rate);
    let metres = qty;
    if (profile) {
      const bl = boyLength(profile);
      if (row.unit === "boy") metres = qty * bl;
      else if (row.unit === "koli") metres = qty * profile.koliMetraj;
    }
    metres = kesin(metres);
    let unitText = `${fmtQty(metres)} mt`;
    if (profile && metres > 0) {
      const kb = koliBoyText(metres, profile);
      if (kb) unitText = `${fmtQty(metres)} mt (${kb})`;
    }
    // Model seçilince kutuya otomatik "-" eklenir; renk yazılmadan
    // gönderilirse sondaki tire ürün adına taşınmasın.
    return satirBitir(
      row,
      row.code.trim().toUpperCase().replace(/-+$/, ""),
      unitText,
      unitPriceTL,
      metres
    );
  }

  if (row.kind === "glass") {
    const sizes = GLASS_SIZES[row.glassType] || [];
    const size = sizes[row.sizeIndex] || sizes[0];
    const plaka = sayi(row.plakaAdet) || 0;
    const price = sayi(row.m2Price) || 0;
    if (!size || plaka <= 0 || price <= 0) return null;
    const m2PerPlaka = plateM2(size);
    const totalM2 = kesin(plaka * m2PerPlaka);
    // Müze camı EUR fiyatlı (₺ seçilirse elle TL yazılır), diğerleri TL
    const priceTL =
      row.glassType === "muze" && row.glassFx !== "tl"
        ? kesin(price * euroRate)
        : price;
    const typeName =
      GLASS_TYPES.find((g) => g.key === row.glassType)?.name || "Cam";
    return satirBitir(
      row,
      typeName,
      // Faturalanan miktar başta: miktar × birim fiyat = satır tutarı
      `${fmtQty(totalM2)} m² · ${plaka} plaka × ${fmtQty(m2PerPlaka)} (${size.label})`,
      priceTL,
      totalM2
    );
  }

  if (row.kind === "ayna") {
    const size = AYNA_SIZES[row.sizeIndex] || AYNA_SIZES[0];
    const plaka = sayi(row.plakaAdet) || 0;
    const price = sayi(row.m2Price) || 0;
    if (plaka <= 0 || price <= 0) return null;
    const m2PerPlaka = plateM2(size);
    const totalM2 = kesin(plaka * m2PerPlaka);
    return satirBitir(
      row,
      "Ayna",
      `${fmtQty(totalM2)} m² · ${plaka} plaka × ${fmtQty(m2PerPlaka)} (${size.label})`,
      price,
      totalM2
    );
  }

  if (row.kind === "technical") {
    const product = getTechnicalProduct(row.techCode);
    const kutu = sayi(row.kutuAdet) || 0;
    if (!product || kutu <= 0) return null;
    const manual = sayi(row.kutuPrice) || 0;
    let kutuPriceTL = 0;
    let priceInfo = "";
    if (product.priceTL != null) {
      kutuPriceTL = manual > 0 ? manual : product.priceTL;
      priceInfo = `₺${fmtPrice(kutuPriceTL)}/kutu`;
    } else {
      const eur = manual > 0 ? manual : product.priceEUR || 0;
      kutuPriceTL = kesin(eur * euroRate);
      priceInfo = `€${fmtPrice(eur)}/kutu`;
    }
    const fullName = row.kartonKodu.trim()
      ? `${product.name} (${row.kartonKodu.trim()})`
      : product.name;
    const totalAdet = kutu * product.adetPerKutu;
    return satirBitir(
      row,
      fullName,
      `${kutu} kutu × ${product.adetPerKutu} = ${totalAdet} adt (${priceInfo})`,
      kutuPriceTL,
      kutu
    );
  }

  // other
  const qty = sayi(row.otherQty) || 0;
  const price = sayi(row.otherPrice) || 0;
  if (!row.name.trim() || qty <= 0) return null;
  return satirBitir(row, row.name.trim(), `${fmtQty(qty)} adt`, price, qty);
}

/** Mükerrer sipariş uyarısında gösterilen özet kayıt. */
interface RecentOrder {
  orderId: string;
  dateKey: string;
  employee: string;
  net: number;
  lines?: unknown[];
}

/** "bugün" / "dün" / "3 gün önce" — uyarıyı okunur kılan gün etiketi. */
function gunEtiketi(dateKey: string): string {
  const bugun = new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" });
  if (dateKey === bugun) return "bugün";
  const fark = Math.round(
    (new Date(bugun).getTime() - new Date(dateKey).getTime()) / 86400000
  );
  if (fark === 1) return "dün";
  if (fark > 1) return `${fark} gün önce`;
  return dateKey;
}

export interface InitialOrder {
  dateKey: string;
  orderId: string;
  customer: string;
  note: string;
  rate: number;
  euroRate: number;
  discountPct: number;
  vatApplied: boolean;
  rows?: Partial<Row>[];
}

export default function OrderForm({
  employeeName,
  initialOrder,
}: {
  employeeName: string;
  initialOrder?: InitialOrder;
}) {
  const [rows, setRows] = useState<Row[]>(() => {
    if (initialOrder?.rows?.length) {
      return initialOrder.rows.map((r) => ({ ...emptyRow(), ...r, id: rowSeq++ }));
    }
    return [emptyRow()];
  });
  const [customer, setCustomer] = useState(initialOrder?.customer ?? "");
  // Müşteri defterinden seçildiyse kaydı sipariş kaydına da bağlarız (cari takip)
  const [customerId, setCustomerId] = useState("");
  // Siparişin şubesi — müşteri defterden seçilince kartındaki şube önerilir,
  // personel gerekirse değiştirir (iki şubede de çalışılabiliyor).
  const [branch, setBranch] = useState<"ankara" | "istanbul">("ankara");
  // Sipariş onay SMS'i — varsayılan açık; müşteri defterden seçilmediyse veya
  // telefonu yoksa sunucu sessizce atlar. Düzenleme modunda gönderilmez.
  const [sendSms, setSendSms] = useState(true);
  const [importOpen, setImportOpen] = useState(false);

  /** Yapay zekanın çözümlediği satırları forma ekler. */
  function applyParsed(data: {
    lines: ParsedLine[];
    customer: string;
    note: string;
  }) {
    const yeni: Row[] = data.lines.map((l) => {
      const r = emptyRow();
      if (l.kind === "frame") {
        r.kind = "frame";
        r.code = l.code;
        r.unit = l.unit === "koli" || l.unit === "boy" ? l.unit : "metre";
        r.qty = String(l.qty);
        const p = findProfile(l.code);
        if (p) r.usd = String(p.priceUSD);
      } else if (l.kind === "glass" || l.kind === "ayna") {
        r.kind = l.kind;
        // Cam türünü metinden yakala (mat / müze / düz)
        const t = `${l.code} ${l.note}`.toLowerCase();
        r.glassType = t.includes("mat") ? "mat" : t.includes("müze") || t.includes("muze") ? "muze" : "duz";
        r.plakaAdet = String(l.qty);
      } else if (l.kind === "technical") {
        r.kind = "technical";
        // Yapay zeka ürün adı da döndürebilir: önce koda, sonra ada bakılır
        const t = getTechnicalProduct(l.code) || findTechnicalByName(l.code);
        if (t) {
          r.techCode = t.code;
          r.kutuPrice = String(t.priceTL ?? t.priceEUR ?? "");
        }
        r.kutuAdet = String(l.qty);
      } else {
        r.kind = "other";
        r.name = [l.code, l.note].filter(Boolean).join(" — ");
        r.otherQty = String(l.qty);
      }
      return r;
    });
    if (yeni.length === 0) return;

    setRows((rs) => {
      // Tamamen boş duran ilk satırı ez, doldurulmuş satırları koru
      const dolu = rs.filter(
        (r) => r.code || r.name || r.techCode || r.qty || r.plakaAdet || r.kutuAdet
      );
      return [...dolu, ...yeni];
    });
    if (data.customer && !customer.trim()) setCustomer(data.customer);
    if (data.note && !note.trim()) setNote(data.note);
    setImportOpen(false);
  }
  const [note, setNote] = useState(initialOrder?.note ?? "");
  const [rate, setRate] = useState(
    initialOrder?.rate ? String(initialOrder.rate) : ""
  );
  const [euroRate, setEuroRate] = useState(
    initialOrder?.euroRate ? String(initialOrder.euroRate) : ""
  );
  const [discountPct, setDiscountPct] = useState(
    initialOrder?.discountPct ? String(initialOrder.discountPct) : ""
  );
  const [vat, setVat] = useState(initialOrder?.vatApplied ?? false);
  const [sending, setSending] = useState(false);
  const [ratesAuto, setRatesAuto] = useState(false);
  // Günün kuru yetkili tarafından belirlendiyse çalışanlarda alan kilitlenir
  const [kurKilitli, setKurKilitli] = useState<{ by: string; at: string } | null>(null);
  const [kurYetkilisi, setKurYetkilisi] = useState(false);
  const [result, setResult] = useState<{
    ok: boolean;
    msg: string;
    waLink?: string;
  } | null>(null);

  // ---- Mükerrer sipariş kontrolü ----
  // Aynı müşteriye başka bir çalışan yakın zamanda sipariş girdiyse formda
  // uyarı çıkar; kaydetmeden önce de onay istenir. Defterden seçilen müşteride
  // customerId, elle yazılanda ad üzerinden eşleşir.
  const [sonSiparisler, setSonSiparisler] = useState<RecentOrder[]>([]);
  useEffect(() => {
    if (initialOrder) return; // düzenleme modunda gereksiz
    const ad = customer.trim();
    if (!customerId && ad.length < 3) {
      setSonSiparisler([]);
      return;
    }
    const t = setTimeout(() => {
      const qs = customerId
        ? `musteri=${encodeURIComponent(customerId)}`
        : `musteriAd=${encodeURIComponent(ad)}`;
      fetch(`/api/orders?${qs}&gun=7`)
        .then((r) => (r.ok ? r.json() : null))
        .then((d) => setSonSiparisler(d?.ok ? d.orders || [] : []))
        .catch(() => setSonSiparisler([]));
    }, 500);
    return () => clearTimeout(t);
  }, [customer, customerId, initialOrder]);

  // Günün kuru daha önce girildiyse formu otomatik doldur.
  // Kur yetkili (firma sahibi) tarafından belirlendiyse alan kilitlenir —
  // herkes aynı kurdan sipariş girsin diye.
  useEffect(() => {
    fetch("/api/rates")
      .then((r) => r.json())
      .then((d) => {
        if (!d.ok) return;
        if (d.yetkili) setKurYetkilisi(true);
        if (d.rates) {
          let used = false;
          if (d.rates.rate > 0) {
            setRate((prev) => {
              if (prev) return prev;
              used = true;
              return String(d.rates.rate);
            });
          }
          if (d.rates.euroRate > 0) {
            setEuroRate((prev) => {
              if (prev) return prev;
              used = true;
              return String(d.rates.euroRate);
            });
          }
          if (used) setRatesAuto(true);
          // Düzenlemede sipariş kendi kurunu korur — kilit uygulanmaz
          if (d.rates.sabit && !d.yetkili && !initialOrder) {
            setKurKilitli({ by: d.rates.by, at: d.rates.updatedAt });
            // Kilitliyse formdaki değer her hâlükârda günün kuru olsun
            if (d.rates.rate > 0) setRate(String(d.rates.rate));
            if (d.rates.euroRate > 0) setEuroRate(String(d.rates.euroRate));
          }
        }
      })
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const rateNum = sayi(rate) || 0;
  const euroNum = sayi(euroRate) || 0;

  const lines = rows
    .map((r) => computeRow(r, rateNum, euroNum))
    .filter((l): l is ComputedLine => l !== null);

  const gross = r2(lines.reduce((s, l) => s + l.lineTotal, 0));
  const pct = Math.max(0, sayi(discountPct) || 0) / 100;
  const discount = r2(gross * pct);
  const afterDiscount = r2(Math.max(0, gross - discount));
  const vatAmount = vat ? r2(afterDiscount * 0.2) : 0;
  const net = r2(afterDiscount + vatAmount);

  function update(id: number, patch: Partial<Row>) {
    setRows((rs) => rs.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  }

  async function submit() {
    setResult(null);
    if (!customer.trim()) {
      setResult({ ok: false, msg: "Müşteri adı gerekli." });
      return;
    }
    if (!lines.length) {
      setResult({ ok: false, msg: "En az bir geçerli satır girin." });
      return;
    }
    // Aynı müşteriye yakın zamanda sipariş varsa son bir onay iste —
    // iki çalışanın habersiz aynı siparişi girmesini engeller.
    if (!initialOrder && sonSiparisler.length > 0) {
      const liste = sonSiparisler
        .map((o) => `• ${o.orderId} — ${gunEtiketi(o.dateKey)} — ${o.employee} — ₺ ${fmt(o.net)}`)
        .join("\n");
      const onay = confirm(
        `Bu müşteriye son 7 günde ${sonSiparisler.length} sipariş girilmiş:\n\n${liste}\n\n` +
          "Aynı sipariş ikinci kez girilmiş olabilir. Yine de kaydedilsin mi?"
      );
      if (!onay) return;
    }
    setSending(true);
    try {
      const payload = {
        customer: customer.trim(),
        customerId,
        branch,
        note: note.trim(),
        rate: rateNum,
        euroRate: euroNum,
        discountPct: sayi(discountPct) || 0,
        vatApplied: vat,
        sendSms: initialOrder ? false : sendSms,
        lines,
        rows,
        gross,
        discount,
        vatAmount,
        net,
      };
      const url = initialOrder
        ? `/api/orders/one?d=${initialOrder.dateKey}&id=${encodeURIComponent(initialOrder.orderId)}`
        : "/api/orders";
      const res = await fetch(url, {
        method: initialOrder ? "PUT" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        setResult({ ok: false, msg: data.error || "Sipariş gönderilemedi." });
        return;
      }
      if (initialOrder) {
        setResult({
          ok: true,
          msg: `Sipariş ${initialOrder.orderId} güncellendi. Yeni toplam: ₺ ${fmt(data.net)}`,
        });
      } else {
        const smsMsg = data.smsSent
          ? "Müşteriye SMS gönderildi."
          : data.smsInfo
            ? `SMS gönderilmedi: ${data.smsInfo}`
            : "";
        setResult({
          ok: true,
          msg: [
            `Sipariş ${data.orderId} oluşturuldu.`,
            data.emailSent ? "E-posta gönderildi." : "",
            data.waSent ? "WhatsApp mesajı gönderildi." : "",
            smsMsg,
          ]
            .filter(Boolean)
            .join(" "),
          waLink: data.waLink,
        });
        setRows([emptyRow()]);
        setCustomer("");
        setNote("");
        setDiscountPct("");
        setVat(false);
      }
    } catch {
      setResult({ ok: false, msg: "Sunucu hatası." });
    } finally {
      setSending(false);
    }
  }

  return (
    <div className="card">
      <div className="grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))" }}>
        <div>
          <label>Çalışan</label>
          <input value={employeeName} disabled />
        </div>
        <div>
          <label>Müşteri *</label>
          <CustomerPicker
            value={customer}
            onChange={(v) => {
              setCustomer(v);
              setCustomerId("");
            }}
            onPick={(c) => {
              setCustomerId(c.id);
              if (c.branch === "ankara" || c.branch === "istanbul") {
                setBranch(c.branch);
              }
            }}
          />
        </div>
        <div>
          <label>Şube</label>
          <select
            value={branch}
            onChange={(e) => setBranch(e.target.value as "ankara" | "istanbul")}
          >
            <option value="ankara">Ankara</option>
            <option value="istanbul">İstanbul</option>
          </select>
        </div>
        <div>
          <label>Dolar Kuru (TL/USD)</label>
          <input
            type="text" inputMode="decimal"
            value={rate}
            onChange={(e) => setRate(e.target.value)}
            placeholder="örn. 45"
            disabled={!!kurKilitli}
            title={kurKilitli ? "Günün kuru yetkili tarafından belirlendi" : undefined}
          />
        </div>
        <div>
          <label>Euro Kuru (TL/EUR)</label>
          <input
            type="text" inputMode="decimal"
            value={euroRate}
            onChange={(e) => setEuroRate(e.target.value)}
            placeholder="örn. 48"
            disabled={!!kurKilitli}
            title={kurKilitli ? "Günün kuru yetkili tarafından belirlendi" : undefined}
          />
        </div>
      </div>
      {kurKilitli ? (
        <p style={{ color: "var(--muted)", fontSize: 12, marginTop: 6 }}>
          🔒 Günün kuru <b>{kurKilitli.by}</b> tarafından belirlendi (
          {new Date(kurKilitli.at).toLocaleTimeString("tr-TR", {
            hour: "2-digit",
            minute: "2-digit",
            timeZone: "Europe/Istanbul",
          })}
          ) — siparişler bu kurdan girilir.
        </p>
      ) : (
        <>
          {ratesAuto && (
            <p style={{ color: "var(--muted)", fontSize: 12, marginTop: 6 }}>
              💡 Bugün için girilen kur otomatik yüklendi — gerekirse
              değiştirebilirsiniz.
            </p>
          )}
          {kurYetkilisi && (
            <p style={{ color: "var(--muted)", fontSize: 12, marginTop: 6 }}>
              💱 Günün kurunu{" "}
              <a href="/panel/kur" style={{ fontWeight: 700 }}>
                Günlük Kur
              </a>{" "}
              ekranından belirlerseniz tüm çalışanlar aynı kurdan sipariş girer.
            </p>
          )}
        </>
      )}

      {/* Mükerrer sipariş uyarısı — aynı müşteriye başka bir çalışan
          yakın zamanda sipariş girdiyse burada görünür. */}
      {!initialOrder && sonSiparisler.length > 0 && (
        <div className="notice warn" style={{ marginTop: 12 }}>
          <b>⚠️ Dikkat: bu müşteriye son 7 günde {sonSiparisler.length} sipariş girilmiş.</b>
          <div style={{ marginTop: 8, display: "grid", gap: 4, fontSize: 13 }}>
            {sonSiparisler.map((o) => (
              <div key={o.orderId}>
                <a
                  href={`/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                  target="_blank"
                  rel="noreferrer"
                  style={{ fontWeight: 700 }}
                >
                  {o.orderId}
                </a>{" "}
                · {gunEtiketi(o.dateKey)} · {o.employee} · ₺ {fmt(o.net)}
                {o.lines?.length ? ` · ${o.lines.length} kalem` : ""}
              </div>
            ))}
          </div>
          <div style={{ marginTop: 8, fontSize: 12.5 }}>
            Aynı siparişin ikinci kez girilmediğinden emin olun — numaraya
            tıklayıp içeriğini kontrol edebilirsiniz.
          </div>
        </div>
      )}

      <h2>Sipariş Satırları</h2>
      {rows.map((row) => {
        const computed = computeRow(row, rateNum, euroNum);
        const profile = row.kind === "frame" ? findProfile(row.code) : undefined;
        const tech =
          row.kind === "technical" ? getTechnicalProduct(row.techCode) : undefined;
        const glassSizes =
          row.kind === "glass" ? GLASS_SIZES[row.glassType] || [] : [];

        return (
          <div
            key={row.id}
            style={{
              border: "1px solid var(--border)",
              borderRadius: 12,
              padding: 14,
              marginBottom: 12,
              background: "rgba(255,255,255,0.02)",
            }}
          >
            <div
              className="grid"
              style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))" }}
            >
              <div>
                <label>Tür</label>
                <select
                  value={row.kind}
                  onChange={(e) => update(row.id, { kind: e.target.value as Kind, sizeIndex: 0 })}
                >
                  <option value="frame">Çerçeve Profili</option>
                  <option value="glass">Cam</option>
                  <option value="ayna">Ayna</option>
                  <option value="technical">Teknik Malzeme</option>
                  <option value="other">Diğer</option>
                </select>
              </div>

              {row.kind === "frame" && (
                <>
                  <div>
                    <label>Profil Kodu</label>
                    <input
                      list={`profiles-${row.id}`}
                      value={row.code}
                      onChange={(e) => {
                        const v = e.target.value;
                        const pr = findProfile(v);
                        // Model tam seçildiğinde depo formatına çevirip sona
                        // otomatik "-" ekle: "4501 S" → "4501S-"; renk kodu
                        // aynı kutuya devam yazılır → "4501S-1242".
                        const duz = (s: string) =>
                          s.toUpperCase().replace(/\s+/g, "");
                        const tamSecim =
                          pr && duz(v) === duz(pr.code) && !v.includes("-");
                        update(row.id, {
                          code: tamSecim ? `${duz(pr.code)}-` : v,
                          usd: pr ? String(pr.priceUSD) : row.usd,
                        });
                      }}
                      placeholder="örn. 4501S-1242"
                    />
                    <datalist id={`profiles-${row.id}`}>
                      {FRAME_PROFILES.map((f) => (
                        <option key={f.code} value={f.code}>
                          {f.series} Serisi — ${f.priceUSD}/mt
                        </option>
                      ))}
                    </datalist>
                  </div>
                  <div>
                    <label>Birim</label>
                    <select
                      value={row.unit}
                      onChange={(e) =>
                        update(row.id, { unit: e.target.value as Row["unit"] })
                      }
                    >
                      <option value="metre">Metre</option>
                      <option value="boy">Boy</option>
                      <option value="koli">Koli</option>
                    </select>
                  </div>
                  <div>
                    <label>Miktar</label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.qty}
                      onChange={(e) => update(row.id, { qty: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>{row.fx === "tl" ? "TL/mt (elle)" : "USD/mt"}</label>
                    {/* Tek fiyat kutusu + para birimi seçici. Varsayılan $:
                        liste fiyatı × kur. Müşteriyle yuvarlak TL anlaşılırsa
                        (47,85 → 47) ₺ seçilir, aynı kutuya TL yazılır. */}
                    <div className="fx-wrap">
                      <select
                        className="fx-sel"
                        value={row.fx}
                        onChange={(e) =>
                          update(row.id, { fx: e.target.value as Row["fx"] })
                        }
                        title="Fiyat para birimi — ₺ seçilirse kur yerine yazdığınız TL geçer"
                      >
                        <option value="usd">$</option>
                        <option value="tl">₺</option>
                      </select>
                      {row.fx === "tl" ? (
                        <input
                          type="text" inputMode="decimal"
                          value={row.tl}
                          onChange={(e) => update(row.id, { tl: e.target.value })}
                          placeholder={
                            rateNum > 0 &&
                            (sayi(row.usd) || profile?.priceUSD)
                              ? `oto ₺${fmtPrice(
                                  kesin(
                                    (sayi(row.usd) ||
                                      profile?.priceUSD ||
                                      0) * rateNum
                                  )
                                )}`
                              : "TL fiyat"
                          }
                          title="Elle TL/mt — boş bırakılırsa USD × kur kullanılır"
                        />
                      ) : (
                        <input
                          type="text" inputMode="decimal"
                          value={row.usd}
                          onChange={(e) => update(row.id, { usd: e.target.value })}
                          placeholder={profile ? String(profile.priceUSD) : "USD"}
                        />
                      )}
                    </div>
                  </div>
                </>
              )}

              {row.kind === "glass" && (
                <>
                  <div>
                    <label>Cam Türü</label>
                    <select
                      value={row.glassType}
                      onChange={(e) =>
                        update(row.id, { glassType: e.target.value, sizeIndex: 0 })
                      }
                    >
                      {GLASS_TYPES.map((g) => (
                        <option key={g.key} value={g.key}>
                          {g.name}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label>Plaka Ölçüsü</label>
                    <select
                      value={row.sizeIndex}
                      onChange={(e) =>
                        update(row.id, { sizeIndex: Number(e.target.value) })
                      }
                    >
                      {glassSizes.map((s, i) => (
                        <option key={s.label} value={i}>
                          {s.label}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label>Plaka Adet</label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.plakaAdet}
                      onChange={(e) => update(row.id, { plakaAdet: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>
                      m² Fiyatı (
                      {row.glassType === "muze"
                        ? row.glassFx === "tl"
                          ? "TL, elle"
                          : "EUR"
                        : "TL"}
                      )
                    </label>
                    {row.glassType === "muze" ? (
                      // Müze camı: çerçevedeki $/₺ gibi €/₺ seçilebilir —
                      // ₺'de m² fiyatı kur hesabı olmadan doğrudan yazılır.
                      <div className="fx-wrap">
                        <select
                          className="fx-sel"
                          value={row.glassFx}
                          onChange={(e) =>
                            update(row.id, {
                              glassFx: e.target.value as Row["glassFx"],
                            })
                          }
                          title="Fiyat para birimi — ₺ seçilirse euro kuru yerine yazdığınız TL geçer"
                        >
                          <option value="eur">€</option>
                          <option value="tl">₺</option>
                        </select>
                        <input
                          type="text" inputMode="decimal"
                          value={row.m2Price}
                          onChange={(e) =>
                            update(row.id, { m2Price: e.target.value })
                          }
                          placeholder={row.glassFx === "tl" ? "TL/m²" : "EUR/m²"}
                        />
                      </div>
                    ) : (
                      <input
                        type="text" inputMode="decimal"
                        value={row.m2Price}
                        onChange={(e) => update(row.id, { m2Price: e.target.value })}
                      />
                    )}
                  </div>
                </>
              )}

              {row.kind === "ayna" && (
                <>
                  <div>
                    <label>Plaka Ölçüsü</label>
                    <select
                      value={row.sizeIndex}
                      onChange={(e) =>
                        update(row.id, { sizeIndex: Number(e.target.value) })
                      }
                    >
                      {AYNA_SIZES.map((s, i) => (
                        <option key={s.label} value={i}>
                          {s.label}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label>Plaka Adet</label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.plakaAdet}
                      onChange={(e) => update(row.id, { plakaAdet: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>m² Fiyatı (TL)</label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.m2Price}
                      onChange={(e) => update(row.id, { m2Price: e.target.value })}
                    />
                  </div>
                </>
              )}

              {row.kind === "technical" && (
                <>
                  <div style={{ minWidth: 230 }}>
                    <label>Ürün</label>
                    {/* 117 ürünlük açılır listede aşağıya inmek zordu —
                        aranabilir seçici kullanılıyor. */}
                    <TechnicalPicker
                      value={row.techCode}
                      onPick={(t) =>
                        update(row.id, {
                          techCode: t.code,
                          // Fiyat alanı boşsa listedeki fiyatla dolsun
                          kutuPrice:
                            row.kutuPrice ||
                            String(t.priceTL ?? t.priceEUR ?? ""),
                        })
                      }
                    />
                  </div>
                  {tech?.isKarton && (
                    <div>
                      <label>Karton Kodu</label>
                      <input
                        value={row.kartonKodu}
                        onChange={(e) => update(row.id, { kartonKodu: e.target.value })}
                        placeholder="örn. 107"
                      />
                    </div>
                  )}
                  <div>
                    <label>Kutu Adet</label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.kutuAdet}
                      onChange={(e) => update(row.id, { kutuAdet: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>
                      Kutu Fiyatı ({tech?.priceTL != null ? "TL" : "EUR"})
                    </label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.kutuPrice}
                      onChange={(e) => update(row.id, { kutuPrice: e.target.value })}
                      placeholder={
                        tech
                          ? String(tech.priceTL ?? tech.priceEUR ?? "")
                          : "Fiyat"
                      }
                    />
                  </div>
                </>
              )}

              {row.kind === "other" && (
                <>
                  <div>
                    <label>Ürün Adı</label>
                    <input
                      value={row.name}
                      onChange={(e) => update(row.id, { name: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>Adet</label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.otherQty}
                      onChange={(e) => update(row.id, { otherQty: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>Birim Fiyat (TL)</label>
                    <input
                      type="text" inputMode="decimal"
                      value={row.otherPrice}
                      onChange={(e) => update(row.id, { otherPrice: e.target.value })}
                    />
                  </div>
                </>
              )}

              {/* Satır iskontosu — her türde geçerli; oran satırdan satıra
                  değişebilir (çerçeve %10, teknik %5 gibi). */}
              <div>
                <label>İskonto %</label>
                <input
                  type="text" inputMode="decimal"
                  min="0"
                  max="100"
                  value={row.iskonto}
                  onChange={(e) => update(row.id, { iskonto: e.target.value })}
                  placeholder="0"
                  title="Bu satıra özel indirim — birim fiyata yansır, fişte oran ürünün yanında görünür"
                />
              </div>
            </div>

            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginTop: 10,
              }}
            >
              <span style={{ color: "var(--text-2)", fontSize: 13 }}>
                {row.kind === "frame" && profile && (
                  <span style={{ color: "var(--brand-light)", marginRight: 10 }}>
                    {profile.code}: 1 koli = {profile.koliAdet} adet /{" "}
                    {profile.koliMetraj.toLocaleString("tr-TR")} mt · 1 boy ={" "}
                    {boyLength(profile).toLocaleString("tr-TR", {
                      maximumFractionDigits: 2,
                    })}{" "}
                    mt
                  </span>
                )}
                {computed
                  ? `${computed.unitText} — Tutar: ₺ ${fmt(computed.lineTotal)}`
                  : "Satır henüz eksik"}
              </span>
              <button
                className="btn small danger"
                onClick={() => setRows((rs) => rs.filter((r) => r.id !== row.id))}
                disabled={rows.length === 1}
              >
                Satırı Sil
              </button>
            </div>
          </div>
        );
      })}

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <button className="btn secondary" onClick={() => setRows((rs) => [...rs, emptyRow()])}>
          + Satır Ekle
        </button>
        <button className="btn secondary" onClick={() => setImportOpen(true)}>
          🤖 Metinden Sipariş Oluştur
        </button>
      </div>

      {importOpen && (
        <OrderTextImport
          onClose={() => setImportOpen(false)}
          onApply={applyParsed}
        />
      )}

      <h2>Özet</h2>
      <div className="grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" }}>
        <div>
          <label>İskonto (%)</label>
          <input
            type="text" inputMode="decimal"
            value={discountPct}
            onChange={(e) => setDiscountPct(e.target.value)}
          />
        </div>
        <div>
          <label>KDV</label>
          <select
            value={vat ? "1" : "0"}
            onChange={(e) => setVat(e.target.value === "1")}
          >
            <option value="0">KDV Yok</option>
            <option value="1">KDV %20</option>
          </select>
        </div>
        <div>
          <label>Not</label>
          <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Sipariş notu" />
        </div>
      </div>

      <table style={{ marginTop: 16, maxWidth: 420 }}>
        <tbody>
          <tr>
            <td>Ara Toplam</td>
            <td style={{ textAlign: "right" }}>₺ {fmt(gross)}</td>
          </tr>
          <tr>
            <td>İskonto</td>
            <td style={{ textAlign: "right" }}>₺ {fmt(discount)}</td>
          </tr>
          <tr>
            <td>KDV</td>
            <td style={{ textAlign: "right" }}>{vat ? `₺ ${fmt(vatAmount)}` : "—"}</td>
          </tr>
          <tr>
            <td>
              <strong>Genel Toplam</strong>
            </td>
            <td style={{ textAlign: "right" }}>
              <strong style={{ color: "var(--brand-light)" }}>₺ {fmt(net)}</strong>
            </td>
          </tr>
        </tbody>
      </table>

      {result && (
        <div className={`notice ${result.ok ? "ok" : "err"}`}>
          {result.msg}
          {result.waLink && (
            <>
              {" "}
              <a href={result.waLink} target="_blank" rel="noreferrer">
                → WhatsApp&apos;tan gönder
              </a>
            </>
          )}
        </div>
      )}

      {!initialOrder && (
        <label
          style={{
            marginTop: 16,
            display: "flex",
            gap: 8,
            alignItems: "center",
            cursor: "pointer",
            fontSize: 14,
          }}
        >
          <input
            type="checkbox"
            checked={sendSms}
            onChange={(e) => setSendSms(e.target.checked)}
            style={{ width: "auto", margin: 0 }}
          />
          Müşteriye &quot;siparişiniz alınmıştır&quot; SMS&apos;i gönder
          {!customerId && sendSms && (
            <span style={{ color: "var(--muted)", fontSize: 13 }}>
              (müşteri defterden seçilirse gönderilir)
            </span>
          )}
        </label>
      )}

      <div style={{ marginTop: 16, display: "flex", gap: 12 }}>
        <button className="btn" onClick={submit} disabled={sending}>
          {sending
            ? "Gönderiliyor…"
            : initialOrder
              ? "Değişiklikleri Kaydet"
              : "Siparişi Gönder"}
        </button>
      </div>
    </div>
  );
}
