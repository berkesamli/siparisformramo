"use client";

/* eslint-disable @next/next/no-img-element */

// Perakende çerçeveletme sihirbazı — eski "Olga Çerçeve Hesaplayıcı"
// (Google Apps Script) uygulamasının Next.js portu.
// Akış: Ölçüler → Çerçeve → Paspartu → Cam → Baskı → Özet → Müşteri & Gönder

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import {
  MAT_TYPES,
  INNER_MAT_TYPES,
  GLASS_TYPES,
  PRINT_TYPES,
  PASPARTU_COLORS,
  computeRetailCosts,
  toMM,
  type MatType,
  type GlassType,
  type PrintType,
} from "@/data/perakende";
import { findFrameImage } from "@/data/frame-images";
import FramePreview from "@/components/FramePreview";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

interface ColorSel {
  code: string;
  hex: string;
}

interface WizardItem {
  artWidth: number;
  artWidthUnit: "cm" | "mm";
  artHeight: number;
  artHeightUnit: "cm" | "mm";
  frameCode: string;
  framePriceTL: number;
  manualPrice: boolean;
  matType: string;
  matCode: string;
  matColor: string;
  matColorHex: string;
  doubleMat: boolean;
  innerMatType: string;
  innerMatColor: string;
  innerMatColorHex: string;
  altMontaj: string;
  zeminEnabled: boolean;
  zeminType: string;
  zeminColor: string;
  zeminColorHex: string;
  matTop: number;
  matRight: number;
  matBottom: number;
  matLeft: number;
  glassType: string;
  printType: string;
  frameCost: number;
  matCost: number;
  glassCost: number;
  printCost: number;
  itemTotal: number;
}

const STEPS = ["Ölçüler", "Çerçeve", "Paspartu", "Cam", "Baskı", "Özet", "Müşteri"];

// Yaygın eser boyutları (cm)
const SIZE_PRESETS: { label: string; w: number; h: number }[] = [
  { label: "A4", w: 21, h: 29.7 },
  { label: "A3", w: 29.7, h: 42 },
  { label: "A2", w: 42, h: 59.4 },
  { label: "A1", w: 59.4, h: 84.1 },
  { label: "A0", w: 84.1, h: 118.9 },
  { label: "15×21", w: 15, h: 21 },
  { label: "30×40", w: 30, h: 40 },
  { label: "50×50", w: 50, h: 50 },
  { label: "50×70", w: 50, h: 70 },
];

function itemShortText(it: WizardItem): string {
  const parts = [
    `${it.artWidth}${it.artWidthUnit} x ${it.artHeight}${it.artHeightUnit}`,
    `Çerçeve: ${it.frameCode}`,
  ];
  if (it.matType !== "Paspartu Yok") {
    let m = `Paspartu: ${it.matType}`;
    if (it.matColor && it.matColor !== "-") m += ` ${it.matColor}`;
    if (it.doubleMat) m += ` + İç: ${it.innerMatType} ${it.innerMatColor}`;
    if (it.zeminEnabled) m += ` | Zemin: ${it.zeminType} ${it.zeminColor}`;
    parts.push(m);
  }
  if (it.glassType !== "Cam Yok") parts.push(`Cam: ${it.glassType}`);
  if (it.printType !== "Baskı Yok") parts.push(`Baskı: ${it.printType}`);
  return parts.join(" | ");
}

function normalizePhoneWa(phone: string): string {
  const digits = String(phone || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("90") && digits.length === 12) return digits;
  if (digits.startsWith("0")) return "9" + digits;
  if (digits.startsWith("5") && digits.length === 10) return "90" + digits;
  return digits;
}

export default function RetailWizard({ employeeName }: { employeeName: string }) {
  const [step, setStep] = useState(1);

  // ---- Müşteri (son adım) ----
  const [customerName, setCustomerName] = useState("");
  const [customerPhone, setCustomerPhone] = useState("");
  const [customerEmail, setCustomerEmail] = useState("");
  const [deliveryDate, setDeliveryDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() + 7);
    return d.toISOString().split("T")[0];
  });
  const [usdRate, setUsdRate] = useState("");
  const [rateAuto, setRateAuto] = useState(false);

  // ---- Ölçüler ----
  const [artWidth, setArtWidth] = useState("");
  const [artHeight, setArtHeight] = useState("");
  const [wUnit, setWUnit] = useState<"cm" | "mm">("cm");
  const [hUnit, setHUnit] = useState<"cm" | "mm">("cm");
  const [imageUrl, setImageUrl] = useState<string | null>(null);

  // ---- Çerçeve ----
  const [seriesCode, setSeriesCode] = useState("");
  const [colorCode, setColorCode] = useState("");
  const [manualPrice, setManualPrice] = useState("");
  const [lookup, setLookup] = useState<{
    found: boolean;
    tlPerM: number;
    resolvedCode?: string;
  } | null>(null);
  const [looking, setLooking] = useState(false);

  // ---- Paspartu ----
  const [mat, setMat] = useState<MatType>(MAT_TYPES[0]);
  const [doubleMat, setDoubleMat] = useState(false);
  const [outerColor, setOuterColor] = useState<ColorSel>({ code: "", hex: "" });
  const [innerMat, setInnerMat] = useState<MatType>(INNER_MAT_TYPES[0]);
  const [innerColor, setInnerColor] = useState<ColorSel>({ code: "", hex: "" });
  const [altMontaj, setAltMontaj] = useState("5");
  const [zeminEnabled, setZeminEnabled] = useState(false);
  const [zeminMat, setZeminMat] = useState<MatType>(INNER_MAT_TYPES[0]);
  const [zeminColor, setZeminColor] = useState<ColorSel>({ code: "", hex: "" });
  const [mTop, setMTop] = useState("0");
  const [mRight, setMRight] = useState("0");
  const [mBottom, setMBottom] = useState("0");
  const [mLeft, setMLeft] = useState("0");

  // ---- Cam & Baskı ----
  const [glass, setGlass] = useState<GlassType>(GLASS_TYPES[0]);
  const [print, setPrint] = useState<PrintType>(PRINT_TYPES[0]);

  // ---- Sepet & sonuç ----
  const [cart, setCart] = useState<WizardItem[]>([]);
  const [discountValue, setDiscountValue] = useState("0");
  const [discountType, setDiscountType] = useState<"percent" | "tl">("percent");
  const [notes, setNotes] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [successId, setSuccessId] = useState<string | null>(null);
  const [error, setError] = useState("");

  // Günlük kuru otomatik doldur (sipariş panelindeki kur kaydından)
  useEffect(() => {
    fetch("/api/rates")
      .then((r) => (r.ok ? r.json() : null))
      .then((d) => {
        if (d?.rates?.rate > 0) {
          setUsdRate(String(d.rates.rate));
          setRateAuto(true);
        }
      })
      .catch(() => {});
  }, []);

  // Seri kodu -> perakende metre fiyatı (sunucudan)
  const lookupTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    const code = seriesCode.trim();
    const rate = parseFloat(usdRate) || 0;
    setLookup(null);
    if (!code || !(rate > 0)) return;
    setLooking(true);
    if (lookupTimer.current) clearTimeout(lookupTimer.current);
    lookupTimer.current = setTimeout(() => {
      fetch(
        `/api/perakende/frame-price?code=${encodeURIComponent(code)}&rate=${rate}`
      )
        .then((r) => (r.ok ? r.json() : null))
        .then((d) => {
          setLooking(false);
          if (d) setLookup(d);
        })
        .catch(() => setLooking(false));
    }, 350);
  }, [seriesCode, usdRate]);

  const framePriceTL = useMemo(() => {
    const mp = parseFloat(manualPrice);
    if (mp > 0) return mp;
    return lookup?.found ? lookup.tlPerM : 0;
  }, [manualPrice, lookup]);

  const wMM = toMM(parseFloat(artWidth) || 0, wUnit);
  const hMM = toMM(parseFloat(artHeight) || 0, hUnit);
  const edges = {
    top: parseFloat(mTop) || 0,
    right: parseFloat(mRight) || 0,
    bottom: parseFloat(mBottom) || 0,
    left: parseFloat(mLeft) || 0,
  };

  const costs = useMemo(
    () =>
      computeRetailCosts({
        wMM,
        hMM,
        matTop: edges.top,
        matRight: edges.right,
        matBottom: edges.bottom,
        matLeft: edges.left,
        framePriceTL,
        matPrice: mat.price,
        doubleMat,
        innerMatPrice: innerMat.price,
        zeminEnabled,
        zeminPrice: zeminMat.price,
        glassPrice: glass.price,
        printUsdPerM2: print.usdPerM2,
        usdRate: parseFloat(usdRate) || 0,
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [
      wMM, hMM, edges.top, edges.right, edges.bottom, edges.left,
      framePriceTL, mat, doubleMat, innerMat, zeminEnabled, zeminMat,
      glass, print, usdRate,
    ]
  );

  const cartTotal = cart.reduce((s, it) => s + it.itemTotal, 0);
  const currentCounts = wMM > 0 && hMM > 0;
  const gross = cartTotal + (currentCounts ? costs.itemTotal : 0);
  const discount = useMemo(() => {
    const v = parseFloat(discountValue) || 0;
    if (v <= 0) return 0;
    const d = discountType === "percent" ? gross * (v / 100) : v;
    return Math.min(d, gross);
  }, [discountValue, discountType, gross]);
  const grandTotal = gross - discount;

  const fullFrameCode =
    (seriesCode.trim().toUpperCase() || "") +
    (colorCode.trim() ? "-" + colorCode.trim().toUpperCase() : "");

  // Gerçek çerçeve görseli (SKU eşleşirse border olarak kullanılır)
  const frameImg = useMemo(() => findFrameImage(fullFrameCode), [fullFrameCode]);
  const bareFrame = Boolean(frameImg?.bareFrame);

  function selectMat(m: MatType) {
    setMat(m);
    setOuterColor({ code: "", hex: "" });
    if (m.price > 0) {
      // Paspartu seçilince kenarlar 0 ise 50 mm varsayılanı
      if (!(parseFloat(mTop) > 0)) setMTop("50");
      if (!(parseFloat(mRight) > 0)) setMRight("50");
      if (!(parseFloat(mBottom) > 0)) setMBottom("50");
      if (!(parseFloat(mLeft) > 0)) setMLeft("50");
    } else {
      setDoubleMat(false);
      setZeminEnabled(false);
      setMTop("0");
      setMRight("0");
      setMBottom("0");
      setMLeft("0");
    }
  }

  function handleImage(file: File | undefined) {
    if (!file || !file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = (e) => setImageUrl(String(e.target?.result || ""));
    reader.readAsDataURL(file);
  }

  const captureItem = useCallback((): WizardItem => {
    return {
      artWidth: parseFloat(artWidth) || 0,
      artWidthUnit: wUnit,
      artHeight: parseFloat(artHeight) || 0,
      artHeightUnit: hUnit,
      frameCode: fullFrameCode || "OZEL",
      framePriceTL,
      manualPrice: parseFloat(manualPrice) > 0,
      matType: mat.name,
      matCode: mat.code,
      matColor: outerColor.code || "-",
      matColorHex: outerColor.hex || "-",
      doubleMat: doubleMat && mat.price > 0,
      innerMatType: doubleMat ? innerMat.name : "-",
      innerMatColor: doubleMat ? innerColor.code || "-" : "-",
      innerMatColorHex: doubleMat ? innerColor.hex || "-" : "-",
      altMontaj: doubleMat ? altMontaj || "5" : "-",
      zeminEnabled: zeminEnabled && mat.price > 0,
      zeminType: zeminEnabled ? zeminMat.name : "-",
      zeminColor: zeminEnabled ? zeminColor.code || "-" : "-",
      zeminColorHex: zeminEnabled ? zeminColor.hex || "-" : "-",
      matTop: edges.top,
      matRight: edges.right,
      matBottom: edges.bottom,
      matLeft: edges.left,
      glassType: glass.name,
      printType: print.name,
      frameCost: Math.round(costs.frameCost * 100) / 100,
      matCost: Math.round(costs.matCost * 100) / 100,
      glassCost: Math.round(costs.glassCost * 100) / 100,
      printCost: Math.round(costs.printCost * 100) / 100,
      itemTotal: Math.round(costs.itemTotal * 100) / 100,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    artWidth, artHeight, wUnit, hUnit, fullFrameCode, framePriceTL, manualPrice,
    mat, outerColor, doubleMat, innerMat, innerColor, altMontaj,
    zeminEnabled, zeminMat, zeminColor, edges.top, edges.right, edges.bottom,
    edges.left, glass, print, costs,
  ]);

  function resetProduct() {
    setArtWidth("");
    setArtHeight("");
    setImageUrl(null);
    setSeriesCode("");
    setColorCode("");
    setManualPrice("");
    setLookup(null);
    setMat(MAT_TYPES[0]);
    setDoubleMat(false);
    setOuterColor({ code: "", hex: "" });
    setInnerMat(INNER_MAT_TYPES[0]);
    setInnerColor({ code: "", hex: "" });
    setAltMontaj("5");
    setZeminEnabled(false);
    setZeminMat(INNER_MAT_TYPES[0]);
    setZeminColor({ code: "", hex: "" });
    setMTop("0");
    setMRight("0");
    setMBottom("0");
    setMLeft("0");
    setGlass(GLASS_TYPES[0]);
    setPrint(PRINT_TYPES[0]);
  }

  function addToCart() {
    if (!currentCounts) {
      setError("Ürün ölçüleri eksik.");
      return;
    }
    setCart([...cart, captureItem()]);
    resetProduct();
    setStep(1);
    setError("");
  }

  function validateStep(s: number): boolean {
    setError("");
    if (s === 1 && !currentCounts) {
      setError("Lütfen eser ölçülerini girin.");
      return false;
    }
    if (s === 2 && framePriceTL <= 0) {
      setError(
        "Çerçeve fiyatı yok — geçerli bir seri kodu ve USD kuru girin veya manuel metre fiyatı yazın."
      );
      return false;
    }
    if (s === 5 && print.usdPerM2 > 0 && !(parseFloat(usdRate) > 0)) {
      setError("Baskı fiyatı için USD kuru gerekli (Çerçeve adımında girin).");
      return false;
    }
    if (s === 7 && (!customerName.trim() || !customerPhone.trim())) {
      setError("Lütfen müşteri adı ve telefon girin.");
      return false;
    }
    return true;
  }

  function next() {
    if (!validateStep(step)) return;
    setStep(Math.min(step + 1, 7));
  }

  async function submitOrder() {
    if (!currentCounts) {
      setError("Lütfen eser ölçülerini girin.");
      setStep(1);
      return;
    }
    if (framePriceTL <= 0) {
      setError("Çerçeve fiyatı eksik — seri kodu veya manuel fiyat girin.");
      setStep(2);
      return;
    }
    if (!customerName.trim() || !customerPhone.trim()) {
      setError("Lütfen müşteri adı ve telefon girin.");
      return;
    }
    setSubmitting(true);
    setError("");
    try {
      const items = [...cart, captureItem()];
      const res = await fetch("/api/perakende/orders", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          customerName: customerName.trim(),
          customerPhone: customerPhone.trim(),
          customerEmail: customerEmail.trim(),
          usdRate: parseFloat(usdRate) || 0,
          deliveryDate,
          notes: notes.trim(),
          items,
          discount: Math.round(discount * 100) / 100,
        }),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Sipariş kaydedilemedi");
      setSuccessId(d.orderId);
    } catch (e: any) {
      setError(e.message || "Bir hata oluştu");
    } finally {
      setSubmitting(false);
    }
  }

  function resetAll() {
    setSuccessId(null);
    setCustomerName("");
    setCustomerPhone("");
    setCustomerEmail("");
    setNotes("");
    setDiscountValue("0");
    setDiscountType("percent");
    setCart([]);
    resetProduct();
    setStep(1);
  }

  function sendWhatsAppQuote() {
    const phone = normalizePhoneWa(customerPhone);
    if (!phone) {
      setError("WhatsApp teklifi için müşteri telefonu gerekli.");
      return;
    }
    const items = currentCounts ? [...cart, captureItem()] : [...cart];
    const lines = ["*Olga Çerçeve — Fiyat Teklifi*", ""];
    items.forEach((it, i) => {
      lines.push(`${i + 1}) ${itemShortText(it)}`);
      lines.push(`   Tutar: ${fmt(it.itemTotal)} TL`);
    });
    lines.push("");
    if (discount > 0) lines.push(`İndirim: -${fmt(discount)} TL`);
    lines.push(`*GENEL TOPLAM: ${fmt(grandTotal)} TL*`);
    lines.push("");
    lines.push("Olga Çerçeve | 0850 305 75 45 | www.olgacerceve.com");
    window.open(
      `https://wa.me/${phone}?text=${encodeURIComponent(lines.join("\n"))}`,
      "_blank"
    );
  }

  const palette = (price: number) => PASPARTU_COLORS[price] || [];

  function ColorPalette({
    price,
    selected,
    onSelect,
  }: {
    price: number;
    selected: ColorSel;
    onSelect: (c: ColorSel) => void;
  }) {
    return (
      <div className="rw-palette">
        {palette(price).map(([code, hex, metallic]) => (
          <button
            key={code}
            type="button"
            title={code}
            className={`rw-color-item ${selected.code === code ? "sel" : ""}`}
            onClick={() => onSelect({ code, hex })}
          >
            <span
              className={`rw-color ${selected.code === code ? "sel" : ""} ${metallic ? "metallic" : ""}`}
              style={{ background: hex }}
            />
            <span className="rw-color-code">{code}</span>
          </button>
        ))}
        {selected.code && (
          <span className="rw-color-label">Seçili: {selected.code}</span>
        )}
      </div>
    );
  }

  if (successId) {
    return (
      <div className="card" style={{ maxWidth: 560, margin: "40px auto", textAlign: "center" }}>
        <div style={{ fontSize: 52 }}>✅</div>
        <h2>Sipariş Kaydedildi</h2>
        <p style={{ fontSize: 15 }}>
          Sipariş Numarası:{" "}
          <strong style={{ color: "var(--brand)", fontSize: 20 }}>{successId}</strong>
        </p>
        <div style={{ display: "flex", gap: 10, justifyContent: "center", marginTop: 16, flexWrap: "wrap" }}>
          <button className="btn" onClick={resetAll}>Yeni Sipariş</button>
          <Link href="/panel/perakende/siparisler" className="btn secondary">
            Perakende Siparişler
          </Link>
        </div>
      </div>
    );
  }

  return (
    <div className="rw-layout">
      <div>
        {/* Adım noktaları */}
        <div className="rw-steps">
          {STEPS.map((label, i) => {
            const n = i + 1;
            return (
              <button
                key={label}
                type="button"
                className={`rw-step ${step === n ? "active" : ""} ${step > n ? "done" : ""}`}
                onClick={() => {
                  if (n < step || validateStep(step)) setStep(n);
                }}
              >
                <span className="rw-step-no">{step > n ? "✓" : n}</span>
                <span className="rw-step-label">{label}</span>
              </button>
            );
          })}
        </div>

        {cart.length > 0 && (
          <div className="notice info" style={{ marginBottom: 14 }}>
            🛒 Sepette <strong>{cart.length}</strong> ürün var — şu an{" "}
            <strong>{cart.length + 1}. ürünü</strong> giriyorsunuz. Toplam: ₺{fmt(cartTotal)}
          </div>
        )}

        {/* 1 — ÖLÇÜLER */}
        {step === 1 && (
          <div className="card">
            <h2 style={{ marginTop: 0 }}>📐 Eser Ölçüleri</h2>
            <div className="rw-grid2">
              <div>
                <label>Genişlik</label>
                <div style={{ display: "flex", gap: 8 }}>
                  <input type="number" min="0" value={artWidth} onChange={(e) => setArtWidth(e.target.value)} placeholder="örn. 50" />
                  <select style={{ width: 84 }} value={wUnit} onChange={(e) => setWUnit(e.target.value as "cm" | "mm")}>
                    <option value="cm">cm</option>
                    <option value="mm">mm</option>
                  </select>
                </div>
              </div>
              <div>
                <label>Yükseklik</label>
                <div style={{ display: "flex", gap: 8 }}>
                  <input type="number" min="0" value={artHeight} onChange={(e) => setArtHeight(e.target.value)} placeholder="örn. 70" />
                  <select style={{ width: 84 }} value={hUnit} onChange={(e) => setHUnit(e.target.value as "cm" | "mm")}>
                    <option value="cm">cm</option>
                    <option value="mm">mm</option>
                  </select>
                </div>
              </div>
            </div>

            <div style={{ marginTop: 16 }}>
              <label>Yaygın Boyutlar</label>
              <div className="rw-presets">
                {SIZE_PRESETS.map((s) => {
                  const active =
                    wUnit === "cm" && hUnit === "cm" &&
                    parseFloat(artWidth) === s.w && parseFloat(artHeight) === s.h;
                  return (
                    <button
                      key={s.label}
                      type="button"
                      className={`rw-preset ${active ? "sel" : ""}`}
                      onClick={() => {
                        setWUnit("cm");
                        setHUnit("cm");
                        setArtWidth(String(s.w));
                        setArtHeight(String(s.h));
                      }}
                    >
                      <strong>{s.label}</strong>
                      <span>{s.w}×{s.h} cm</span>
                    </button>
                  );
                })}
              </div>
            </div>

            <div style={{ marginTop: 18 }}>
              <label>Eser Görseli (opsiyonel — önizlemede gösterilir)</label>
              {imageUrl ? (
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <img src={imageUrl} alt="Eser" style={{ maxHeight: 110, borderRadius: 8, border: "1px solid var(--border)" }} />
                  <button className="btn small secondary" onClick={() => setImageUrl(null)}>Kaldır</button>
                </div>
              ) : (
                <label className="rw-upload">
                  <input
                    type="file"
                    accept="image/*"
                    style={{ display: "none" }}
                    onChange={(e) => handleImage(e.target.files?.[0])}
                  />
                  🖼️ Fotoğraf seçmek için tıklayın
                </label>
              )}
            </div>
          </div>
        )}

        {/* 2 — ÇERÇEVE */}
        {step === 2 && (
          <div className="card">
            <h2 style={{ marginTop: 0 }}>🖼️ Çerçeve Seçimi</h2>
            <div className="rw-grid2">
              <div>
                <label>Seri Kodu</label>
                <input
                  value={seriesCode}
                  onChange={(e) => setSeriesCode(e.target.value.toUpperCase())}
                  placeholder="örn. GC065"
                />
                <span
                  style={{
                    fontSize: 12.5,
                    fontWeight: 600,
                    color: looking
                      ? "var(--muted)"
                      : lookup?.found
                        ? "var(--success)"
                        : seriesCode
                          ? "var(--error)"
                          : "var(--muted)",
                  }}
                >
                  {looking
                    ? "Aranıyor..."
                    : lookup?.found
                      ? `✓ ${lookup.resolvedCode} — ₺${fmt(lookup.tlPerM)}/metre`
                      : seriesCode
                        ? parseFloat(usdRate) > 0
                          ? "✗ Kod bulunamadı — manuel fiyat girebilirsiniz"
                          : "Önce USD kurunu girin"
                        : "Seri kodunu girin"}
                </span>
              </div>
              <div>
                <label>Renk Kodu (opsiyonel)</label>
                <input
                  value={colorCode}
                  onChange={(e) => setColorCode(e.target.value.toUpperCase())}
                  placeholder="örn. 1473"
                />
                {frameImg && (
                  <span style={{ fontSize: 12, fontWeight: 600, color: "var(--success)" }}>
                    ✓ {frameImg.sku} — gerçek görseli önizlemede
                  </span>
                )}
              </div>
              <div>
                <label>USD Kuru (TL)</label>
                <input
                  type="number"
                  step="0.01"
                  value={usdRate}
                  onChange={(e) => { setUsdRate(e.target.value); setRateAuto(false); }}
                  placeholder="örn. 47.50"
                />
                {rateAuto && (
                  <span style={{ fontSize: 12, color: "var(--success)" }}>
                    ✓ Bugünün kuru otomatik geldi
                  </span>
                )}
              </div>
              <div>
                <label>Manuel Metre Fiyatı (₺/m)</label>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={manualPrice}
                  onChange={(e) => setManualPrice(e.target.value)}
                  placeholder="Kod yoksa elle girin"
                />
              </div>
            </div>
            {framePriceTL > 0 && (
              <div className="notice ok" style={{ marginTop: 14 }}>
                Çerçeve metre fiyatı: <strong>₺{fmt(framePriceTL)}/m</strong>
                {parseFloat(manualPrice) > 0 ? " (manuel)" : ""}
              </div>
            )}
          </div>
        )}

        {/* 3 — PASPARTU */}
        {step === 3 && (
          <div className="card">
            <h2 style={{ marginTop: 0 }}>🎨 Paspartu</h2>
            <div className="rw-options">
              {MAT_TYPES.map((m) => (
                <button
                  key={m.code}
                  type="button"
                  className={`rw-option ${mat.code === m.code ? "sel" : ""}`}
                  onClick={() => selectMat(m)}
                >
                  <span className="rw-option-icon">{m.icon}</span>
                  <span className="rw-option-name">{m.name}</span>
                </button>
              ))}
            </div>

            {mat.price > 0 && (
              <>
                <div style={{ marginTop: 18 }}>
                  <label>Paspartu Katı</label>
                  <div className="rw-toggle">
                    <button type="button" className={!doubleMat ? "active" : ""} onClick={() => setDoubleMat(false)}>Tek</button>
                    <button type="button" className={doubleMat ? "active" : ""} onClick={() => setDoubleMat(true)}>Çift</button>
                  </div>
                </div>

                <div style={{ marginTop: 16 }}>
                  <label>{doubleMat ? "Dış Paspartu Rengi" : "Paspartu Rengi"} — {mat.name}</label>
                  <ColorPalette price={mat.price} selected={outerColor} onSelect={setOuterColor} />
                </div>

                {doubleMat && (
                  <>
                    <div style={{ marginTop: 16 }}>
                      <label>İç Paspartu Türü</label>
                      <div className="rw-minis">
                        {INNER_MAT_TYPES.map((m) => (
                          <button
                            key={m.code}
                            type="button"
                            className={`rw-mini ${innerMat.code === m.code ? "sel" : ""}`}
                            onClick={() => { setInnerMat(m); setInnerColor({ code: "", hex: "" }); }}
                          >
                            {m.name}
                          </button>
                        ))}
                      </div>
                      <div style={{ marginTop: 10 }}>
                        <ColorPalette price={innerMat.price} selected={innerColor} onSelect={setInnerColor} />
                      </div>
                    </div>
                    <div style={{ marginTop: 14, maxWidth: 220 }}>
                      <label>Alt Montaj (mm)</label>
                      <input type="number" min="0" value={altMontaj} onChange={(e) => setAltMontaj(e.target.value)} />
                    </div>
                  </>
                )}

                <div style={{ marginTop: 18 }}>
                  <label>Zemin (arka fon paspartusu)</label>
                  <div className="rw-toggle">
                    <button type="button" className={!zeminEnabled ? "active" : ""} onClick={() => setZeminEnabled(false)}>Zemin Yok</button>
                    <button type="button" className={zeminEnabled ? "active" : ""} onClick={() => setZeminEnabled(true)}>Zemin Var</button>
                  </div>
                  {zeminEnabled && (
                    <div style={{ marginTop: 10 }}>
                      <div className="rw-minis">
                        {INNER_MAT_TYPES.map((m) => (
                          <button
                            key={m.code}
                            type="button"
                            className={`rw-mini ${zeminMat.code === m.code ? "sel" : ""}`}
                            onClick={() => { setZeminMat(m); setZeminColor({ code: "", hex: "" }); }}
                          >
                            {m.name}
                          </button>
                        ))}
                      </div>
                      <div style={{ marginTop: 10 }}>
                        <ColorPalette price={zeminMat.price} selected={zeminColor} onSelect={setZeminColor} />
                      </div>
                    </div>
                  )}
                </div>

                <div style={{ marginTop: 18 }}>
                  <label>Kenar Ölçüleri (mm)</label>
                  <div className="rw-grid4">
                    <div><span className="rw-edge-label">Üst</span><input type="number" min="0" value={mTop} onChange={(e) => setMTop(e.target.value)} /></div>
                    <div><span className="rw-edge-label">Sağ</span><input type="number" min="0" value={mRight} onChange={(e) => setMRight(e.target.value)} /></div>
                    <div><span className="rw-edge-label">Alt</span><input type="number" min="0" value={mBottom} onChange={(e) => setMBottom(e.target.value)} /></div>
                    <div><span className="rw-edge-label">Sol</span><input type="number" min="0" value={mLeft} onChange={(e) => setMLeft(e.target.value)} /></div>
                  </div>
                </div>
              </>
            )}
          </div>
        )}

        {/* 4 — CAM */}
        {step === 4 && (
          <div className="card">
            <h2 style={{ marginTop: 0 }}>🪟 Cam Seçimi</h2>
            <div className="rw-options">
              {GLASS_TYPES.map((g) => (
                <button
                  key={g.name}
                  type="button"
                  className={`rw-option ${glass.name === g.name ? "sel" : ""}`}
                  onClick={() => setGlass(g)}
                >
                  <span className="rw-option-icon">{g.icon}</span>
                  <span className="rw-option-name">{g.name}</span>
                  <span className="rw-option-desc">{g.desc}</span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 5 — BASKI */}
        {step === 5 && (
          <div className="card">
            <h2 style={{ marginTop: 0 }}>🖨️ Baskı (opsiyonel)</h2>
            <p className="subtitle" style={{ marginTop: 0 }}>
              Baskı, eserin kendi ölçüsü üzerinden hesaplanır.
            </p>
            <div className="rw-options">
              {PRINT_TYPES.map((p) => (
                <button
                  key={p.name}
                  type="button"
                  className={`rw-option ${print.name === p.name ? "sel" : ""}`}
                  onClick={() => setPrint(p)}
                >
                  <span className="rw-option-icon">{p.icon}</span>
                  <span className="rw-option-name">{p.name}</span>
                  <span className="rw-option-desc">{p.desc}</span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 6 — ÖZET */}
        {step === 6 && (
          <div className="card">
            <h2 style={{ marginTop: 0 }}>
              🧾 {cart.length > 0 ? `Şu Anki Ürün (${cart.length + 1}. Ürün)` : "Sipariş Özeti"}
            </h2>
            <table style={{ marginBottom: 14 }}>
              <tbody>
                <tr><td>Ölçü</td><td>{artWidth || "-"} {wUnit} × {artHeight || "-"} {hUnit}</td></tr>
                <tr><td>Çerçeve</td><td>{fullFrameCode || "-"} {framePriceTL > 0 && `(₺${fmt(framePriceTL)}/m)`}</td></tr>
                <tr>
                  <td>Paspartu</td>
                  <td>
                    {mat.name}
                    {mat.price > 0 && outerColor.code && ` — ${outerColor.code}`}
                    {doubleMat && ` + İç: ${innerMat.name}${innerColor.code ? ` — ${innerColor.code}` : ""}`}
                    {zeminEnabled && ` | Zemin: ${zeminMat.name}${zeminColor.code ? ` — ${zeminColor.code}` : ""}`}
                  </td>
                </tr>
                <tr><td>Cam</td><td>{glass.name}</td></tr>
                <tr><td>Baskı</td><td>{print.name}</td></tr>
              </tbody>
            </table>

            {cart.length > 0 && (
              <div style={{ marginBottom: 14 }}>
                <label>Sepetteki Ürünler</label>
                {cart.map((it, i) => (
                  <div className="rw-cart-item" key={i}>
                    <div>
                      <strong>{i + 1}. Ürün</strong>
                      <div style={{ fontSize: 12.5, color: "var(--text-2)" }}>{itemShortText(it)}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <strong>₺{fmt(it.itemTotal)}</strong>
                      <button
                        className="btn small danger"
                        onClick={() => setCart(cart.filter((_, j) => j !== i))}
                      >
                        Sil
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            <div className="rw-totals">
              <div><span>Çerçeve</span><span>₺{fmt(costs.frameCost)}</span></div>
              {costs.matCost > 0 && <div><span>Paspartu</span><span>₺{fmt(costs.matCost)}</span></div>}
              {costs.glassCost > 0 && <div><span>Cam</span><span>₺{fmt(costs.glassCost)}</span></div>}
              {costs.printCost > 0 && <div><span>Baskı</span><span>₺{fmt(costs.printCost)}</span></div>}
              {cart.length > 0 && (
                <>
                  <div><span>Bu Ürün</span><span>₺{fmt(costs.itemTotal)}</span></div>
                  <div><span>Sepet ({cart.length} ürün)</span><span>₺{fmt(cartTotal)}</span></div>
                </>
              )}
              {discount > 0 && (
                <div style={{ color: "var(--error)" }}><span>İndirim</span><span>-₺{fmt(discount)}</span></div>
              )}
              <div className="rw-grand"><span>GENEL TOPLAM</span><span>₺{fmt(grandTotal)}</span></div>
            </div>

            <div className="rw-grid2" style={{ marginTop: 16 }}>
              <div>
                <label>İndirim</label>
                <div style={{ display: "flex", gap: 8 }}>
                  <input type="number" min="0" value={discountValue} onChange={(e) => setDiscountValue(e.target.value)} />
                  <select style={{ width: 84 }} value={discountType} onChange={(e) => setDiscountType(e.target.value as "percent" | "tl")}>
                    <option value="percent">%</option>
                    <option value="tl">TL</option>
                  </select>
                </div>
              </div>
              <div>
                <label>Sipariş Notu</label>
                <input value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Not (opsiyonel)" />
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, marginTop: 20, flexWrap: "wrap" }}>
              <button className="btn secondary" onClick={addToCart}>
                ➕ Sepete Ekle & Yeni Ürün
              </button>
            </div>
          </div>
        )}

        {/* 7 — MÜŞTERİ & GÖNDER */}
        {step === 7 && (
          <div className="card">
            <h2 style={{ marginTop: 0 }}>👤 Müşteri Bilgileri</h2>
            <div className="rw-grid2">
              <div>
                <label>Ad Soyad *</label>
                <input value={customerName} onChange={(e) => setCustomerName(e.target.value)} placeholder="Müşteri adı" />
              </div>
              <div>
                <label>Telefon *</label>
                <input value={customerPhone} onChange={(e) => setCustomerPhone(e.target.value)} placeholder="05xx xxx xx xx" />
              </div>
              <div>
                <label>E-posta</label>
                <input type="email" value={customerEmail} onChange={(e) => setCustomerEmail(e.target.value)} placeholder="ornek@eposta.com" />
              </div>
              <div>
                <label>Teslim Tarihi</label>
                <input type="date" value={deliveryDate} onChange={(e) => setDeliveryDate(e.target.value)} />
              </div>
            </div>

            <div className="rw-totals" style={{ marginTop: 18 }}>
              <div>
                <span>{cart.length > 0 ? `${cart.length + 1} ürün` : "1 ürün"}</span>
                <span></span>
              </div>
              {discount > 0 && (
                <div style={{ color: "var(--error)" }}><span>İndirim</span><span>-₺{fmt(discount)}</span></div>
              )}
              <div className="rw-grand"><span>GENEL TOPLAM</span><span>₺{fmt(grandTotal)}</span></div>
            </div>

            <div style={{ display: "flex", gap: 10, marginTop: 20, flexWrap: "wrap" }}>
              <button className="btn wa" onClick={sendWhatsAppQuote}>
                📲 WhatsApp Teklif Gönder
              </button>
              <button className="btn" disabled={submitting} onClick={submitOrder}>
                {submitting ? "Kaydediliyor..." : "✅ Siparişi Kaydet ve Gönder"}
              </button>
            </div>
            <p style={{ fontSize: 12, color: "var(--muted)", marginTop: 10 }}>
              Personel: {employeeName}
            </p>
          </div>
        )}

        {error && (
          <div className="notice err" style={{ marginTop: 14 }}>
            {error}
          </div>
        )}

        {/* Alt gezinme */}
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 18 }}>
          <button className="btn secondary" disabled={step === 1} onClick={() => setStep(step - 1)}>
            ← Geri
          </button>
          {step < 7 && (
            <button className="btn" onClick={next}>
              İleri →
            </button>
          )}
        </div>
      </div>

      {/* Sağ — canlı önizleme (web sitesindeki hesaplayıcı tasarımının portu) */}
      <aside className="rw-preview-panel">
        <div className="card rw-preview-card" style={{ position: "sticky", top: 90 }}>
          <FramePreview
            wMM={wMM}
            hMM={hMM}
            matTop={edges.top}
            matRight={edges.right}
            matBottom={edges.bottom}
            matLeft={edges.left}
            matPrice={mat.price}
            matName={mat.name}
            matColorCode={outerColor.code}
            matColorHex={outerColor.hex}
            doubleMat={doubleMat}
            innerMatPrice={innerMat.price}
            innerMatName={innerMat.name}
            innerColorCode={innerColor.code}
            innerColorHex={innerColor.hex}
            mountingMM={parseFloat(altMontaj) || 5}
            zeminEnabled={zeminEnabled}
            zeminColorHex={zeminColor.hex}
            glassName={glass.name}
            frameImg={frameImg}
            fullCode={fullFrameCode}
            artImageUrl={imageUrl}
          />

          {currentCounts && costs.itemTotal > 0 && (
            <div className="fp-cost">
              <div className="fp-cost-bar">
                <span className="seg frame" style={{ width: `${(costs.frameCost / costs.itemTotal) * 100}%` }} />
                <span className="seg mat" style={{ width: `${(costs.matCost / costs.itemTotal) * 100}%` }} />
                <span className="seg glass" style={{ width: `${(costs.glassCost / costs.itemTotal) * 100}%` }} />
                <span className="seg print" style={{ width: `${(costs.printCost / costs.itemTotal) * 100}%` }} />
              </div>
              <div className="fp-cost-legend">
                <span><i className="dot frame" /> Çerçeve <strong>₺{fmt(costs.frameCost)}</strong></span>
                {costs.matCost > 0 && <span><i className="dot mat" /> Paspartu <strong>₺{fmt(costs.matCost)}</strong></span>}
                {costs.glassCost > 0 && <span><i className="dot glass" /> Cam <strong>₺{fmt(costs.glassCost)}</strong></span>}
                {costs.printCost > 0 && <span><i className="dot print" /> Baskı <strong>₺{fmt(costs.printCost)}</strong></span>}
              </div>
            </div>
          )}

          <div className="rw-preview-total">
            <span>Bu ürün</span>
            <strong>₺{fmt(currentCounts ? costs.itemTotal : 0)}</strong>
          </div>
          {cart.length > 0 && (
            <div className="rw-preview-total" style={{ borderTop: "none", paddingTop: 0 }}>
              <span>Genel toplam</span>
              <strong style={{ color: "var(--brand)" }}>₺{fmt(grandTotal)}</strong>
            </div>
          )}
        </div>
      </aside>
    </div>
  );
}
