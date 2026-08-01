"use client";

import { useEffect, useMemo, useState } from "react";
import { FRAME_PROFILES, findProfile, boyLength, koliBoyText } from "@/data/catalog";
import {
  TECHNICAL_PRODUCTS,
  getTechnicalProduct,
  technicalByCategory,
} from "@/data/technical";
import { GLASS_TYPES, GLASS_SIZES, AYNA_SIZES, plateM2 } from "@/data/glass";

type Kind = "frame" | "glass" | "ayna" | "technical" | "other";

interface Row {
  id: number;
  kind: Kind;
  // frame
  code: string;
  unit: "metre" | "boy" | "koli";
  qty: string;
  usd: string;
  // glass / ayna
  glassType: string;
  sizeIndex: number;
  plakaAdet: string;
  m2Price: string;
  // technical
  techCode: string;
  kartonKodu: string;
  kutuAdet: string;
  kutuPrice: string; // TL veya EUR (ürüne göre)
  // other
  name: string;
  otherQty: string;
  otherPrice: string;
}

let rowSeq = 1;
const emptyRow = (): Row => ({
  id: rowSeq++,
  kind: "frame",
  code: "",
  unit: "metre",
  qty: "",
  usd: "",
  glassType: "duz",
  sizeIndex: 0,
  plakaAdet: "",
  m2Price: "",
  techCode: "",
  kartonKodu: "",
  kutuAdet: "",
  kutuPrice: "",
  name: "",
  otherQty: "",
  otherPrice: "",
});

const r2 = (n: number) => Math.round(n * 100) / 100;
const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

interface ComputedLine {
  name: string;
  unitText: string;
  unitPriceTL: number;
  lineTotal: number;
}

function computeRow(row: Row, rate: number, euroRate: number): ComputedLine | null {
  if (row.kind === "frame") {
    const profile = findProfile(row.code);
    const qty = parseFloat(row.qty) || 0;
    if (!row.code.trim() || qty <= 0) return null;
    const usd = parseFloat(row.usd) || profile?.priceUSD || 0;
    const unitPriceTL = r2(usd * rate);
    let metres = qty;
    if (profile) {
      const bl = boyLength(profile);
      if (row.unit === "boy") metres = qty * bl;
      else if (row.unit === "koli") metres = qty * profile.koliMetraj;
    }
    metres = r2(metres);
    let unitText = `${fmt(metres)} mt`;
    if (profile && metres > 0) {
      const kb = koliBoyText(metres, profile);
      if (kb) unitText = `${fmt(metres)} mt (${kb})`;
    }
    return {
      name: row.code.trim().toUpperCase(),
      unitText,
      unitPriceTL,
      lineTotal: r2(metres * unitPriceTL),
    };
  }

  if (row.kind === "glass") {
    const sizes = GLASS_SIZES[row.glassType] || [];
    const size = sizes[row.sizeIndex] || sizes[0];
    const plaka = parseFloat(row.plakaAdet) || 0;
    const price = parseFloat(row.m2Price) || 0;
    if (!size || plaka <= 0 || price <= 0) return null;
    const m2PerPlaka = plateM2(size);
    const totalM2 = r2(plaka * m2PerPlaka);
    // Müze camı EUR fiyatlı, diğerleri TL
    const priceTL = row.glassType === "muze" ? r2(price * euroRate) : price;
    const typeName =
      GLASS_TYPES.find((g) => g.key === row.glassType)?.name || "Cam";
    return {
      name: typeName,
      unitText: `${plaka} plaka × ${fmt(m2PerPlaka)} m² = ${fmt(totalM2)} m² (${size.label})`,
      unitPriceTL: priceTL,
      lineTotal: r2(totalM2 * priceTL),
    };
  }

  if (row.kind === "ayna") {
    const size = AYNA_SIZES[row.sizeIndex] || AYNA_SIZES[0];
    const plaka = parseFloat(row.plakaAdet) || 0;
    const price = parseFloat(row.m2Price) || 0;
    if (plaka <= 0 || price <= 0) return null;
    const m2PerPlaka = plateM2(size);
    const totalM2 = r2(plaka * m2PerPlaka);
    return {
      name: "Ayna",
      unitText: `${plaka} plaka × ${fmt(m2PerPlaka)} m² = ${fmt(totalM2)} m² (${size.label})`,
      unitPriceTL: price,
      lineTotal: r2(totalM2 * price),
    };
  }

  if (row.kind === "technical") {
    const product = getTechnicalProduct(row.techCode);
    const kutu = parseFloat(row.kutuAdet) || 0;
    if (!product || kutu <= 0) return null;
    const manual = parseFloat(row.kutuPrice) || 0;
    let kutuPriceTL = 0;
    let priceInfo = "";
    if (product.priceTL != null) {
      kutuPriceTL = manual > 0 ? manual : product.priceTL;
      priceInfo = `₺${fmt(kutuPriceTL)}/kutu`;
    } else {
      const eur = manual > 0 ? manual : product.priceEUR || 0;
      kutuPriceTL = r2(eur * euroRate);
      priceInfo = `€${fmt(eur)}/kutu`;
    }
    const fullName = row.kartonKodu.trim()
      ? `${product.name} (${row.kartonKodu.trim()})`
      : product.name;
    const totalAdet = kutu * product.adetPerKutu;
    return {
      name: fullName,
      unitText: `${kutu} kutu × ${product.adetPerKutu} = ${totalAdet} adt (${priceInfo})`,
      unitPriceTL: kutuPriceTL,
      lineTotal: r2(kutu * kutuPriceTL),
    };
  }

  // other
  const qty = parseFloat(row.otherQty) || 0;
  const price = parseFloat(row.otherPrice) || 0;
  if (!row.name.trim() || qty <= 0) return null;
  return {
    name: row.name.trim(),
    unitText: `${qty} adt`,
    unitPriceTL: price,
    lineTotal: r2(qty * price),
  };
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
  const [result, setResult] = useState<{
    ok: boolean;
    msg: string;
    waLink?: string;
  } | null>(null);

  // Günün kuru daha önce girildiyse formu otomatik doldur
  useEffect(() => {
    fetch("/api/rates")
      .then((r) => r.json())
      .then((d) => {
        if (d.ok && d.rates) {
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
        }
      })
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const techCategories = useMemo(() => technicalByCategory(), []);
  const rateNum = parseFloat(rate) || 0;
  const euroNum = parseFloat(euroRate) || 0;

  const lines = rows
    .map((r) => computeRow(r, rateNum, euroNum))
    .filter((l): l is ComputedLine => l !== null);

  const gross = r2(lines.reduce((s, l) => s + l.lineTotal, 0));
  const pct = Math.max(0, parseFloat(discountPct) || 0) / 100;
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
    setSending(true);
    try {
      const payload = {
        customer: customer.trim(),
        note: note.trim(),
        rate: rateNum,
        euroRate: euroNum,
        discountPct: parseFloat(discountPct) || 0,
        vatApplied: vat,
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
        setResult({
          ok: true,
          msg: `Sipariş ${data.orderId} oluşturuldu. ${data.emailSent ? "E-posta gönderildi." : "E-posta yapılandırılmadı (SMTP env eksik)."} ${data.waSent ? "WhatsApp mesajı gönderildi." : ""}`,
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
          <input
            value={customer}
            onChange={(e) => setCustomer(e.target.value)}
            placeholder="Müşteri / Firma adı"
          />
        </div>
        <div>
          <label>Dolar Kuru (TL/USD)</label>
          <input
            type="number"
            step="0.01"
            value={rate}
            onChange={(e) => setRate(e.target.value)}
            placeholder="örn. 45"
          />
        </div>
        <div>
          <label>Euro Kuru (TL/EUR)</label>
          <input
            type="number"
            step="0.01"
            value={euroRate}
            onChange={(e) => setEuroRate(e.target.value)}
            placeholder="örn. 48"
          />
        </div>
      </div>
      {ratesAuto && (
        <p style={{ color: "var(--muted)", fontSize: 12, marginTop: 6 }}>
          💡 Bugün için girilen kur otomatik yüklendi — gerekirse
          değiştirebilirsiniz.
        </p>
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
                        const pr = findProfile(e.target.value);
                        update(row.id, {
                          code: e.target.value,
                          usd: pr ? String(pr.priceUSD) : row.usd,
                        });
                      }}
                      placeholder="örn. KS 2030"
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
                      type="number"
                      step="0.1"
                      value={row.qty}
                      onChange={(e) => update(row.id, { qty: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>USD/mt</label>
                    <input
                      type="number"
                      step="0.01"
                      value={row.usd}
                      onChange={(e) => update(row.id, { usd: e.target.value })}
                      placeholder={profile ? String(profile.priceUSD) : "USD"}
                    />
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
                      type="number"
                      value={row.plakaAdet}
                      onChange={(e) => update(row.id, { plakaAdet: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>
                      m² Fiyatı ({row.glassType === "muze" ? "EUR" : "TL"})
                    </label>
                    <input
                      type="number"
                      step="0.01"
                      value={row.m2Price}
                      onChange={(e) => update(row.id, { m2Price: e.target.value })}
                    />
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
                      type="number"
                      value={row.plakaAdet}
                      onChange={(e) => update(row.id, { plakaAdet: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>m² Fiyatı (TL)</label>
                    <input
                      type="number"
                      step="0.01"
                      value={row.m2Price}
                      onChange={(e) => update(row.id, { m2Price: e.target.value })}
                    />
                  </div>
                </>
              )}

              {row.kind === "technical" && (
                <>
                  <div>
                    <label>Ürün</label>
                    <select
                      value={row.techCode}
                      onChange={(e) => update(row.id, { techCode: e.target.value })}
                    >
                      <option value="">Ürün Seçiniz…</option>
                      {Object.entries(techCategories).map(([cat, products]) => (
                        <optgroup key={cat} label={cat}>
                          {products.map((t) => (
                            <option key={t.code} value={t.code}>
                              {t.name}
                            </option>
                          ))}
                        </optgroup>
                      ))}
                    </select>
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
                      type="number"
                      value={row.kutuAdet}
                      onChange={(e) => update(row.id, { kutuAdet: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>
                      Kutu Fiyatı ({tech?.priceTL != null ? "TL" : "EUR"})
                    </label>
                    <input
                      type="number"
                      step="0.01"
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
                      type="number"
                      value={row.otherQty}
                      onChange={(e) => update(row.id, { otherQty: e.target.value })}
                    />
                  </div>
                  <div>
                    <label>Birim Fiyat (TL)</label>
                    <input
                      type="number"
                      step="0.01"
                      value={row.otherPrice}
                      onChange={(e) => update(row.id, { otherPrice: e.target.value })}
                    />
                  </div>
                </>
              )}
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

      <button className="btn secondary" onClick={() => setRows((rs) => [...rs, emptyRow()])}>
        + Satır Ekle
      </button>

      <h2>Özet</h2>
      <div className="grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" }}>
        <div>
          <label>İskonto (%)</label>
          <input
            type="number"
            step="0.5"
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
