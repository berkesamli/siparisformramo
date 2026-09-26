"use client";

import { useEffect, useMemo, useState } from "react";
// Fiyat listesi bilinçli olarak istemci paketinde değil — oturumla
// /api/katalog'dan çekilir (useKatalog); yardımcılar veri içermez.
import {
  boyLength,
  koliBoyText,
  profilBul,
  teknikBul,
  type TechnicalProduct,
} from "@/lib/catalog-utils";
import { useKatalog, type Katalog } from "@/lib/use-katalog";
import { GLASS_TYPES, GLASS_SIZES, AYNA_SIZES, plateM2 } from "@/data/glass";
import { kurus, kesin, fmtQty, fmtPrice, fmtTL, sayi } from "@/lib/num";
import { stokEslesme, toBoy } from "@/lib/stock-search";
import type { StockItem } from "@/lib/stock-parse";
import CustomerPicker from "@/components/CustomerPicker";
import MikroCariKutusu from "@/components/MikroCariKutusu";
import TechnicalPicker from "@/components/TechnicalPicker";
import OrderTextImport, { type ParsedLine, type ParsedResult } from "@/components/OrderTextImport";
import Icon from "@/components/shell/Icon";

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
  // Euro fiyatlı teknik ürünlerde (Scappi, NS karton...) ₺ seçilirse kutu
  // fiyatı kur hesabı olmadan doğrudan TL yazılır.
  techFx: "eur" | "tl";
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
  unit: "koli", // çerçeve profili çoğunlukla koliyle satılır; ilk seçenek koli
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
  techFx: "eur",
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

function findTechnicalByName(list: TechnicalProduct[], q: string) {
  const k = sadeAd(q);
  if (k.length < 3) return undefined;
  return list.find((t) => {
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

function computeRow(
  row: Row,
  rate: number,
  euroRate: number,
  katalog: Katalog
): ComputedLine | null {
  if (row.kind === "frame") {
    const profile = profilBul(katalog.profiles, row.code);
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
    const product = teknikBul(katalog.technical, row.techCode);
    const kutu = sayi(row.kutuAdet) || 0;
    if (!product || kutu <= 0) return null;
    const manual = sayi(row.kutuPrice) || 0;
    let kutuPriceTL = 0;
    let priceInfo = "";
    if (product.priceTL != null) {
      kutuPriceTL = manual > 0 ? manual : product.priceTL;
      priceInfo = `₺${fmtPrice(kutuPriceTL)}/kutu`;
    } else if (row.techFx === "tl" && manual > 0) {
      // Euro fiyatlı üründe ₺ seçildi: kutu fiyatı doğrudan TL
      kutuPriceTL = manual;
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

/** Müşteri kartındaki "Yeni Sipariş" ile gelen ön seçim (/panel?musteri=C123). */
export interface InitialCustomer {
  id: string;
  title: string;
  branch?: "ankara" | "istanbul";
  iskontoPct?: number;
}

export default function OrderForm({
  employeeName,
  initialOrder,
  initialCustomer,
}: {
  employeeName: string;
  initialOrder?: InitialOrder;
  initialCustomer?: InitialCustomer;
}) {
  // Fiyat listesi (çerçeve + teknik) oturumla sunucudan gelir
  const katalog = useKatalog();
  const [rows, setRows] = useState<Row[]>(() => {
    if (initialOrder?.rows?.length) {
      return initialOrder.rows.map((r) => ({ ...emptyRow(), ...r, id: rowSeq++ }));
    }
    return [emptyRow()];
  });
  const [customer, setCustomer] = useState(
    initialOrder?.customer ?? initialCustomer?.title ?? ""
  );
  // Müşteri defterinden seçildiyse kaydı sipariş kaydına da bağlarız (cari takip)
  const [customerId, setCustomerId] = useState(initialCustomer?.id ?? "");
  // Siparişin şubesi — müşteri defterden seçilince kartındaki şube önerilir,
  // personel gerekirse değiştirir (iki şubede de çalışılabiliyor).
  const [branch, setBranch] = useState<"ankara" | "istanbul">(
    initialCustomer?.branch === "istanbul" ? "istanbul" : "ankara"
  );
  // Sipariş onay SMS'i — varsayılan açık; müşteri defterden seçilmediyse veya
  // telefonu yoksa sunucu sessizce atlar. Düzenleme modunda gönderilmez.
  const [sendSms, setSendSms] = useState(true);
  const [importOpen, setImportOpen] = useState(false);

  /** Yapay zekanın çözümlediği satırları forma ekler. */
  function applyParsed(data: ParsedResult) {
    const yeni: Row[] = data.lines.map((l) => {
      const r = emptyRow();
      if (l.kind === "frame") {
        r.kind = "frame";
        r.code = l.code;
        r.unit = l.unit === "koli" || l.unit === "boy" ? l.unit : "metre";
        r.qty = String(l.qty);
        const p = profilBul(katalog.profiles, l.code);
        if (p) r.usd = String(p.priceUSD);
      } else if (l.kind === "glass" || l.kind === "ayna") {
        r.kind = l.kind;
        // Cam türünü metinden yakala (mat / müze / düz)
        const t = `${l.code} ${l.note}`.toLowerCase();
        r.glassType = t.includes("mat") ? "mat" : t.includes("müze") || t.includes("muze") ? "muze" : "duz";
        r.plakaAdet = String(l.qty);
      } else if (l.kind === "technical") {
        r.kind = "technical";
        // Çözümleyici ürün kodunu verir; vermediyse koda, sonra ada bakılır
        const t =
          (l.techCode ? teknikBul(katalog.technical, l.techCode) : undefined) ||
          teknikBul(katalog.technical, l.code) ||
          findTechnicalByName(katalog.technical, l.code);
        if (t) {
          r.techCode = t.code;
          r.kutuPrice = String(t.priceTL ?? t.priceEUR ?? "");
        }
        if (l.kartonKodu) r.kartonKodu = l.kartonKodu;
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
    // Metindeki "%40 isk + KDV" gibi başlıklar: iskonto boşsa yazılır, KDV işareti metne göre ayarlanır
    if (data.iskontoPct !== undefined && data.iskontoPct > 0) setDiscountPct((prev) => prev || String(data.iskontoPct));
    if (data.kdv !== undefined) setVat(data.kdv);
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
    initialOrder?.discountPct
      ? String(initialOrder.discountPct)
      : initialCustomer?.iskontoPct
        ? String(initialCustomer.iskontoPct)
        : ""
  );
  // Yeni siparişte KDV varsayılan olarak açık (faturalı satış çoğunlukta);
  // düzenlemede siparişin kendi değeri korunur. Gerekirse kapatılır.
  const [vat, setVat] = useState(initialOrder?.vatApplied ?? true);
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

  // ---- Canlı stok rozeti (yalnız çerçeve profilleri) ----
  // Depo stok listesi bir kez çekilir; her çerçeve satırında kodun altında
  // "ANK X boy · İST Y boy" görünür. Kod stokta hiç yoksa sarı uyarı çıkar.
  const [stokItems, setStokItems] = useState<StockItem[] | null>(null);
  useEffect(() => {
    fetch("/api/stock")
      .then((r) => (r.ok ? r.json() : null))
      .then((d) => {
        if (d?.ok && Array.isArray(d.data?.items)) setStokItems(d.data.items);
      })
      .catch(() => {});
  }, []);

  // Rozetin yanında eşleşen stok kodu da yazılır: "KS4022-BİG" gibi eksik ya da hatalı yazımda başka
  // rengin/modelin stoku gösterilebilir; kod birebir değilse rozet sarıya döner ve "en yakın kod" belirtilir.
  function stokRozet(code: string): { txt: string; bulundu: boolean; kod?: string; tam?: boolean; yon?: "eksik" | "fazla"; aday?: number; adaylar?: string[] } | null {
    if (!stokItems || !stokItems.length) return null;
    const q = code.trim();
    // Kod yeterince yazılmadan rozet gösterme (model seçilirken gürültü olmasın)
    if (q.replace(/[^A-Za-z0-9]/g, "").length < 6) return null;
    // Yalnızca birebir / önek / içerme eşleşmeleri sayılır — bulanık benzerlik yanlış modelin stokunu göstermesin.
    const e = stokEslesme(stokItems, q);
    if (!e) return { txt: "stokta görünmüyor", bulundu: false };
    // Eksik yazımda birden çok olası kod varsa miktar gösterilmez (yanlış rengin stoku sanılmasın)
    if (!e.tam && e.aday > 1) return { txt: `${e.aday} benzer kod — ${e.yon === "fazla" ? "kodu kontrol edin" : "kodu tamamlayın"}`, bulundu: true, kod: e.item.code, tam: false, yon: e.yon, aday: e.aday, adaylar: e.adaylar };
    return {
      txt: `ANK ${toBoy(e.item.ankaraMt)} boy · İST ${toBoy(e.item.istanbulMt)} boy`,
      bulundu: true,
      kod: e.item.code,
      tam: e.tam,
      yon: e.yon,
      aday: e.aday,
      adaylar: e.adaylar,
    };
  }

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
    .map((r) => computeRow(r, rateNum, euroNum, katalog))
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
    // ---- Anomali perdesi (kaydetmeden önce son kontrol) ----
    // Engelleyiciler: ₺0 fiyatlı satır (çoğunlukla kur girilmeden çerçeve
    // satırı) ve ₺0 toplam. Onaylılar: alışılmadık yüksek iskonto.
    const sifirli = lines.find((l) => !(l.unitPriceTL > 0) || !(l.lineTotal > 0));
    if (sifirli) {
      setResult({
        ok: false,
        msg: `"${sifirli.name}" satırının fiyatı ₺0 görünüyor — kur veya birim fiyat eksik. Kontrol edip tekrar deneyin.`,
      });
      return;
    }
    if (!(net > 0)) {
      setResult({ ok: false, msg: "Sipariş toplamı ₺0 — fiyatları kontrol edin." });
      return;
    }
    const uyarilar: string[] = [];
    const genelIsk = sayi(discountPct);
    if (genelIsk >= 30) uyarilar.push(`Genel iskonto %${fmt(genelIsk)}`);
    for (const r of rows) {
      const p = sayi(r.iskonto);
      if (p >= 30) uyarilar.push(`Bir satırda %${fmt(p)} iskonto var`);
    }
    if (uyarilar.length) {
      const onay = confirm(
        `Dikkat — alışılmadık değerler:\n\n• ${uyarilar.join("\n• ")}\n\nYine de kaydedilsin mi?`
      );
      if (!onay) return;
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
        const smsMsg = data.musteriWa
          ? "Müşteriye WhatsApp ile fiş gönderildi (WhatsApp'ı yoksa SMS gidecek)."
          : data.smsSent
            ? "Müşteriye SMS gönderildi."
            : data.smsInfo
              ? `Bildirim gönderilmedi: ${data.smsInfo}`
              : "";
        setResult({
          ok: true,
          msg: [
            `Sipariş ${data.orderId} oluşturuldu.`,
            data.stored === false
              ? "⚠ Kalıcı depo bağlı değil — sipariş panelde SAKLANAMADI."
              : "",
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
        setVat(true);
      }
    } catch {
      setResult({ ok: false, msg: "Sunucu hatası." });
    } finally {
      setSending(false);
    }
  }

  // Kart başlığı şeridi: açılır listeler kartın dışına taşabilsin diye kart
  // "overflow" (görünür) olunca şeridin üst köşeleri kartın yuvarlağını
  // almalı — yoksa köşeler kartın kenarlığından dışarı taşar.
  const kartBas = {
    borderTopLeftRadius: "inherit",
    borderTopRightRadius: "inherit",
  } as const;

  return (
    <>
      {/* 1) Müşteri, şube, kur — müşteri açılır listesi kartın altından
          taşabilir; sonraki kartların üstünde kalsın diye z-index verilir
          (cam kart backdrop-filter ile kendi katmanını açar). */}
      <div className="card overflow" style={{ zIndex: 3 }}>
        <div className="card-head" style={kartBas}>
          <span className="card-head-icon">
            <Icon name="user" size={16} />
          </span>
          <div>
            <h2>Sipariş Bilgileri</h2>
            <span className="card-head-sub">Müşteri, şube ve günün kuru</span>
          </div>
        </div>

        {/* Müşteri kutusu geniş (açılır listede ad + telefon + şehir sığsın),
            şube dar; telefonda alt alta iner. */}
        <div style={{ display: "flex", flexWrap: "wrap", gap: 16 }}>
          <div style={{ flex: "2 1 260px", minWidth: 0 }}>
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
                // Bayiye özel iskonto: kartta tanımlıysa genel iskonto alanına
                // otomatik yazılır (alan doluysa dokunulmaz).
                if (c.iskontoPct && c.iskontoPct > 0) {
                  setDiscountPct((prev) => prev || String(c.iskontoPct));
                }
              }}
            />
            {/* Kayıtlı müşteri seçilince Mikro'daki resmi bakiye/vade uyarısı;
                eşleşme yoksa buradan tek tıkla bağlanır. */}
            {customerId && <MikroCariKutusu customerId={customerId} compact />}
          </div>
          <div style={{ flex: "1 1 140px", minWidth: 0 }}>
            <label>Şube</label>
            <select
              value={branch}
              onChange={(e) => setBranch(e.target.value as "ankara" | "istanbul")}
            >
              <option value="ankara">Ankara</option>
              <option value="istanbul">İstanbul</option>
            </select>
          </div>
        </div>
        <div
          className="grid"
          style={{
            gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
            marginTop: 16,
          }}
        >
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
          <div>
            <label>Çalışan</label>
            <input value={employeeName} disabled />
          </div>
        </div>
        {kurKilitli ? (
          <p className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>
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
              <p className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>
                💡 Bugün için girilen kur otomatik yüklendi — gerekirse
                değiştirebilirsiniz.
              </p>
            )}
            {kurYetkilisi && (
              <p className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>
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
          <div className="notice warn" style={{ marginTop: 12, marginBottom: 0 }}>
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
      </div>

      {/* 2) Satırlar — teknik malzeme açılır listesi geniş açılır ve kartın
          kenarından taşabilir; overflow serbest + alttaki kartın üstünde. */}
      <div className="card overflow" style={{ zIndex: 2 }}>
        <div className="card-head" style={kartBas}>
          <span className="card-head-icon">
            <Icon name="list" size={16} />
          </span>
          <div>
            <h2>Sipariş Satırları</h2>
            <span className="card-head-sub">{rows.length} satır</span>
          </div>
        </div>
        {rows.map((row) => {
          const computed = computeRow(row, rateNum, euroNum, katalog);
          const profile =
            row.kind === "frame" ? profilBul(katalog.profiles, row.code) : undefined;
          const tech =
            row.kind === "technical"
              ? teknikBul(katalog.technical, row.techCode)
              : undefined;
          const glassSizes =
            row.kind === "glass" ? GLASS_SIZES[row.glassType] || [] : [];

          return (
            <div
              key={row.id}
              style={{
                border: "1px solid var(--border)",
                borderRadius: "var(--radius-sm)",
                padding: 14,
                marginBottom: 12,
                background: "var(--surface-2)",
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
                          const pr = profilBul(katalog.profiles, v);
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
                        placeholder="4501S-1242"
                      />
                      <datalist id={`profiles-${row.id}`}>
                        {katalog.profiles.map((f) => (
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
                        <option value="koli">Koli</option>
                        <option value="boy">Boy</option>
                        <option value="metre">Metre</option>
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
                    {/* Sabit genişlik verilmez — auto-fit ızgarada hücreden
                        taşıp yandaki Karton Kodu kutusunun üstüne biniyordu.
                        Geniş açılır liste .tp-menu'da ayrıca sağlanır. */}
                    <div>
                      <label>Ürün</label>
                      {/* 117 ürünlük açılır listede aşağıya inmek zordu —
                          aranabilir seçici kullanılıyor. */}
                      <TechnicalPicker
                        value={row.techCode}
                        products={katalog.technical}
                        onPick={(t) => {
                          update(row.id, {
                            techCode: t.code,
                            // Fiyat alanı boşsa listedeki fiyatla dolsun
                            kutuPrice:
                              row.kutuPrice ||
                              String(t.priceTL ?? t.priceEUR ?? ""),
                          });
                          // Kartonlu üründe (NS/Scappi) seçimden sonra açılan
                          // "Karton Kodu" alanı tablette ekranın altında
                          // kalabiliyor — görünür yere kaydır ve odaklan ki
                          // kod hemen yazılabilsin.
                          if (t.isKarton) {
                            setTimeout(() => {
                              const el = document.getElementById(
                                `karton-kodu-${row.id}`
                              ) as HTMLInputElement | null;
                              el?.scrollIntoView({ behavior: "smooth", block: "center" });
                              el?.focus({ preventScroll: true });
                            }, 100);
                          }
                        }}
                      />
                    </div>
                    {tech?.isKarton && (
                      <div>
                        <label>Karton Kodu</label>
                        <input
                          id={`karton-kodu-${row.id}`}
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
                        Kutu Fiyatı (
                        {tech?.priceTL != null
                          ? "TL"
                          : row.techFx === "tl"
                            ? "TL, elle"
                            : "EUR"}
                        )
                      </label>
                      {tech && tech.priceTL == null ? (
                        // Euro fiyatlı ürün (Scappi, NS karton...): çerçevedeki
                        // $/₺ gibi €/₺ seçilebilir — ₺'de fiyat kur hesabı
                        // olmadan doğrudan TL yazılır.
                        <div className="fx-wrap">
                          <select
                            className="fx-sel"
                            value={row.techFx}
                            onChange={(e) =>
                              update(row.id, {
                                techFx: e.target.value as Row["techFx"],
                                // Önceki para biriminin fiyatı kalmasın
                                kutuPrice: "",
                              })
                            }
                            title="Fiyat para birimi — ₺ seçilirse euro kuru yerine yazdığınız TL geçer"
                          >
                            <option value="eur">€</option>
                            <option value="tl">₺</option>
                          </select>
                          <input
                            type="text" inputMode="decimal"
                            value={row.kutuPrice}
                            onChange={(e) => update(row.id, { kutuPrice: e.target.value })}
                            placeholder={
                              row.techFx === "tl"
                                ? "TL/kutu"
                                : String(tech.priceEUR ?? "EUR")
                            }
                          />
                        </div>
                      ) : (
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
                      )}
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

              {/* Satır altı: bilgi + stok rozeti + tutar solda, sil sağda.
                  Telefonda sarar — sil düğmesi kendi satırına iner. */}
              <div
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  alignItems: "center",
                  gap: "6px 10px",
                  marginTop: 10,
                }}
              >
                <span
                  style={{
                    color: "var(--text-2)",
                    fontSize: 13,
                    flex: "1 1 240px",
                    minWidth: 0,
                    display: "flex",
                    flexWrap: "wrap",
                    alignItems: "center",
                    gap: "4px 10px",
                  }}
                >
                  {row.kind === "frame" && profile && (
                    <span style={{ color: "var(--brand)" }}>
                      {profile.code}: 1 koli = {profile.koliAdet} adet /{" "}
                      {profile.koliMetraj.toLocaleString("tr-TR")} mt · 1 boy ={" "}
                      {boyLength(profile).toLocaleString("tr-TR", {
                        maximumFractionDigits: 2,
                      })}{" "}
                      mt
                    </span>
                  )}
                  {row.kind === "frame" &&
                    (() => {
                      const rz = stokRozet(row.code);
                      if (!rz) return null;
                      // Birebir eşleşme yeşil; eksik/hatalı yazımla en yakın koda düşüldüyse ya da birden çok
                      // olası kod varsa sarı — hangi koda bakıldığı rozetin yanında yazılır
                      const kesin = rz.bulundu && Boolean(rz.tam);
                      const belirsiz = rz.bulundu && !rz.tam && (rz.aday || 1) > 1;
                      return (
                        <>
                          <span
                            className={`badge ${kesin ? "ok" : "warn"}`}
                            title={
                              !rz.bulundu
                                ? "Bu kod depo stok listesinde bulunamadı — kodu kontrol edin"
                                : kesin
                                  ? `Depo stok listesindeki güncel miktar (1 boy = 2,9 mt) — stok kodu: ${rz.kod}`
                                  : belirsiz
                                    ? rz.yon === "fazla"
                                      ? "Yazdığınız kod birden çok stok kodundan uzun; hangisi olduğunu kontrol edin"
                                      : "Yazdığınız kod birden çok stok koduna uyuyor; miktar için kodu tamamlayın"
                                    : `Yazdığınız kod stokta birebir yok; en yakın kodun (${rz.kod}) miktarı gösteriliyor — kodu kontrol edin`
                            }
                          >
                            <Icon name="package" size={12} />
                            {rz.txt}
                          </span>
                          {rz.bulundu && (
                            <span className={`of-stok-kod ${kesin ? "" : "uyari"}`}>
                              {belirsiz
                                ? <>Olası kodlar: {(rz.adaylar || []).join(", ")}{(rz.aday || 0) > (rz.adaylar || []).length ? "…" : ""}</>
                                : <>{kesin ? "Stok kodu" : "En yakın stok kodu"}: <b>{rz.kod}</b></>}
                            </span>
                          )}
                        </>
                      );
                    })()}
                  <span>
                    {computed ? (
                      <>
                        {computed.unitText} — Tutar:{" "}
                        <b style={{ color: "var(--text)" }}>₺ {fmt(computed.lineTotal)}</b>
                      </>
                    ) : (
                      "Satır henüz eksik"
                    )}
                  </span>
                </span>
                <button
                  className="btn small ghost"
                  style={{ color: "var(--error)", marginLeft: "auto" }}
                  onClick={() => setRows((rs) => rs.filter((r) => r.id !== row.id))}
                  disabled={rows.length === 1}
                >
                  <Icon name="x" size={14} />
                  Satırı Sil
                </button>
              </div>
            </div>
          );
        })}

        <div className="row">
          <button className="btn secondary" onClick={() => setRows((rs) => [...rs, emptyRow()])}>
            <Icon name="plus" size={16} />
            Satır Ekle
          </button>
          <button className="btn secondary" onClick={() => setImportOpen(true)}>
            <Icon name="sparkles" size={16} />
            Metinden Sipariş Oluştur
          </button>
        </div>

        {importOpen && (
          <OrderTextImport
            onClose={() => setImportOpen(false)}
            onApply={applyParsed}
          />
        )}
      </div>

      {/* 3) Özet — iskonto, KDV, not, toplamlar ve gönder */}
      <div className="card">
        <div className="card-head" style={kartBas}>
          <span className="card-head-icon">
            <Icon name="percent" size={16} />
          </span>
          <div>
            <h2>Özet</h2>
            <span className="card-head-sub">
              {lines.length} kalem · ₺ {fmt(net)}
            </span>
          </div>
        </div>
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
              <option value="1">KDV %20</option>
              <option value="0">KDV Yok</option>
            </select>
          </div>
          <div>
            <label>Not <span style={{ color: "var(--muted)", fontWeight: 400, textTransform: "none", letterSpacing: "normal" }}>(iç kullanım — müşteriye giden fişte görünmez)</span></label>
            <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Hazırlayanlar için not: acil, peşin ödeme…" />
          </div>
        </div>

        {/* Toplamlar: iki sütun, dar ekranda daralır — yatay taşma olmaz */}
        <table style={{ marginTop: 16, maxWidth: 420 }}>
          <tbody>
            <tr>
              <td>Ara Toplam</td>
              <td className="num">₺ {fmt(gross)}</td>
            </tr>
            <tr>
              <td>İskonto</td>
              <td className="num">₺ {fmt(discount)}</td>
            </tr>
            <tr>
              <td>KDV</td>
              <td className="num">{vat ? `₺ ${fmt(vatAmount)}` : "—"}</td>
            </tr>
            <tr>
              <td>
                <strong>Genel Toplam</strong>
              </td>
              <td className="num">
                <strong style={{ color: "var(--brand)", fontSize: 16 }}>₺ {fmt(net)}</strong>
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
                <a
                  className="btn small wa"
                  style={{ marginTop: 8 }}
                  href={result.waLink}
                  target="_blank"
                  rel="noreferrer"
                >
                  <Icon name="message" size={14} />
                  WhatsApp&apos;tan gönder
                </a>
              </>
            )}
          </div>
        )}

        {!initialOrder && (
          <label
            style={{
              marginTop: 16,
              marginBottom: 0,
              display: "flex",
              flexWrap: "wrap",
              gap: "4px 8px",
              alignItems: "center",
              cursor: "pointer",
              fontSize: 14,
              fontWeight: 500,
              color: "var(--text)",
              textTransform: "none",
              letterSpacing: "normal",
            }}
          >
            <input
              type="checkbox"
              checked={sendSms}
              onChange={(e) => setSendSms(e.target.checked)}
              style={{ width: "auto", margin: 0 }}
            />
            Müşteriye sipariş bildirimi gönder (WhatsApp ile fiş PDF&apos;i, WhatsApp&apos;ı yoksa SMS)
            {!customerId && sendSms && (
              <span style={{ color: "var(--muted)", fontSize: 13 }}>
                (müşteri defterden seçilirse gönderilir)
              </span>
            )}
          </label>
        )}

        <div style={{ marginTop: 16, display: "flex", gap: 12, flexWrap: "wrap" }}>
          <button className="btn" onClick={submit} disabled={sending}>
            {sending
              ? "Gönderiliyor…"
              : initialOrder
                ? "Değişiklikleri Kaydet"
                : "Siparişi Gönder"}
          </button>
        </div>
      </div>
    </>
  );
}
