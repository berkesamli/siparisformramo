"use client";

// Tahsilat giriş penceresi — sipariş listesinden ("şu siparişe ödeme geldi")
// veya cari sayfasından ("bu müşteriden para geldi") açılır. Tarih, tutar,
// yöntem, şube ve tahsil eden bilgisiyle kalıcı tahsilat kaydı oluşturur.

import { useState } from "react";

export interface TahsilatBaglam {
  customerId?: string;
  customerName: string;
  orderId?: string;
  orderDateKey?: string;
  /** Sipariş bağlamında kalan bakiye — tutar alanına önerilir. */
  kalan?: number;
  /** Önerilen şube (müşteri kartından). */
  branch?: "ankara" | "istanbul";
  /**
   * Elden satış modu: müşteri adı serbestçe yazılabilir (ayaküstü perakende,
   * teknik malzeme satışı gibi kartsız küçük tahsilatlar için).
   */
  serbest?: boolean;
}

const YONTEMLER = [
  ["nakit", "Nakit"],
  ["havale", "Havale / EFT"],
  ["krediKarti", "Kredi Kartı"],
  ["cek", "Çek"],
  ["senet", "Senet"],
  ["diger", "Diğer"],
] as const;

const fmt = (n: number) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export default function TahsilatModal({
  baglam,
  onClose,
  onSaved,
}: {
  baglam: TahsilatBaglam;
  onClose: () => void;
  onSaved: () => void;
}) {
  const bugun = new Date().toLocaleDateString("en-CA", {
    timeZone: "Europe/Istanbul",
  });
  const [dateKey, setDateKey] = useState(bugun);
  const [amount, setAmount] = useState(
    baglam.kalan && baglam.kalan > 0 ? String(baglam.kalan) : ""
  );
  const [method, setMethod] = useState<string>("nakit");
  const [currency, setCurrency] = useState<"TL" | "USD" | "EUR">("TL");
  const [branch, setBranch] = useState<"ankara" | "istanbul">(
    baglam.branch || "ankara"
  );
  const [tahsilEden, setTahsilEden] = useState("");
  const [note, setNote] = useState("");
  const [ad, setAd] = useState(baglam.customerName || "");
  const [saving, setSaving] = useState(false);
  const [err, setErr] = useState("");

  const tutar = parseFloat(amount.replace(",", ".")) || 0;

  async function kaydet() {
    if (tutar <= 0) {
      setErr("Tutar sıfırdan büyük olmalı.");
      return;
    }
    if (baglam.serbest && !ad.trim()) {
      setErr("Müşteri / açıklama alanı boş olamaz.");
      return;
    }
    setSaving(true);
    setErr("");
    try {
      const res = await fetch("/api/finans/tahsilat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dateKey,
          amount: tutar,
          method,
          currency,
          branch,
          customerId: baglam.customerId,
          customerName: baglam.serbest ? ad.trim() : baglam.customerName,
          orderId: baglam.orderId,
          orderDateKey: baglam.orderDateKey,
          tahsilEden: tahsilEden.trim() || undefined,
          note: note.trim() || undefined,
        }),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Kaydedilemedi");
      onSaved();
      onClose();
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Bir hata oluştu");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <h2 style={{ marginTop: 0 }}>
          {baglam.serbest ? "💰 Elden Satış / Tahsilat" : "💰 Tahsilat Gir"}
        </h2>
        {!baglam.serbest && (
          <p className="subtitle" style={{ marginTop: -6 }}>
            {baglam.customerName}
            {baglam.orderId ? ` — ${baglam.orderId}` : ""}
            {baglam.kalan != null && baglam.kalan > 0 && (
              <> · kalan bakiye ₺ {fmt(baglam.kalan)}</>
            )}
          </p>
        )}

        <div className="rw-grid2">
          {baglam.serbest && (
            <div style={{ gridColumn: "1 / -1" }}>
              <label>Müşteri / Açıklama</label>
              <input
                value={ad}
                onChange={(e) => setAd(e.target.value)}
                placeholder="örn. PERAKENDE — çerçeve yapımı, teknik malzeme satışı…"
              />
            </div>
          )}
          <div>
            <label>Tarih</label>
            <input
              type="date"
              value={dateKey}
              onChange={(e) => setDateKey(e.target.value)}
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
            <label>Tutar</label>
            <input
              type="number"
              step="0.01"
              min="0"
              value={amount}
              onChange={(e) => setAmount(e.target.value)}
              placeholder="0,00"
            />
          </div>
          <div>
            <label>Para Birimi</label>
            <select
              value={currency}
              onChange={(e) => setCurrency(e.target.value as "TL" | "USD" | "EUR")}
            >
              <option value="TL">₺ TL</option>
              <option value="USD">$ USD</option>
              <option value="EUR">€ EUR</option>
            </select>
          </div>
          <div>
            <label>Yöntem</label>
            <select value={method} onChange={(e) => setMethod(e.target.value)}>
              {YONTEMLER.map(([k, l]) => (
                <option key={k} value={k}>
                  {l}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label>Tahsil Eden (opsiyonel)</label>
            <input
              value={tahsilEden}
              onChange={(e) => setTahsilEden(e.target.value)}
              placeholder="örn. Alaattin"
            />
          </div>
          <div style={{ gridColumn: "1 / -1" }}>
            <label>Not</label>
            <input
              value={note}
              onChange={(e) => setNote(e.target.value)}
              placeholder="örn. elden teslim, dekont no…"
            />
          </div>
        </div>

        {(method === "cek" || method === "senet") && (
          <p className="notice info" style={{ marginTop: 10 }}>
            Çek/senet cariyi düşürür ancak kasaya <strong>tahsil edildiğinde</strong>{" "}
            girer. Vade ve banka takibi çek/senet ekranından yapılır (yakında).
          </p>
        )}
        {currency !== "TL" && (
          <p className="notice info" style={{ marginTop: 10 }}>
            Döviz tahsilatı sipariş bakiyesini etkilemez; döviz kasasında ayrı
            izlenir.
          </p>
        )}
        {err && <div className="notice err" style={{ marginTop: 10 }}>{err}</div>}

        <div style={{ marginTop: 14, display: "flex", gap: 10 }}>
          <button className="btn" onClick={kaydet} disabled={saving || tutar <= 0}>
            {saving ? "Kaydediliyor…" : `Kaydet (₺ ${fmt(tutar)})`}
          </button>
          <button className="btn secondary" onClick={onClose}>
            Vazgeç
          </button>
        </div>
      </div>
    </div>
  );
}
