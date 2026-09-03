"use client";

// Bayi sipariş listesi — filtre, arama, durum/ödeme güncelleme, PDF, WhatsApp.

import { useCallback, useEffect, useState } from "react";
import { ORDER_STATUSES, PAYMENT_LABELS, type OrderStatus, type PaymentStatus } from "@/data/pricing";
import type { SavedOrder } from "@/lib/orders";
import { eslesir } from "@/lib/search-norm";

type Row = SavedOrder & { trackUrl?: string };

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const STATUS_COLORS: Record<OrderStatus, string> = {
  Beklemede: "#b45309",
  "Hazırlanıyor": "#1d4ed8",
  "Hazır": "#067a55",
  "Teslim Edildi": "#374151",
  "İptal": "#b91c1c",
};

function waDigits(phone: string): string {
  const digits = String(phone || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("90") && digits.length === 12) return digits;
  if (digits.startsWith("0")) return "9" + digits;
  if (digits.startsWith("5") && digits.length === 10) return "90" + digits;
  return digits;
}

export default function OrdersList({ dealerName }: { dealerName: string }) {
  const [range, setRange] = useState<"today" | "week" | "month" | "all" | "date">("week");
  const [date, setDate] = useState("");
  const [query, setQuery] = useState("");
  const [orders, setOrders] = useState<Row[]>([]);
  const [loading, setLoading] = useState(true);
  const [blob, setBlob] = useState(true);
  const [open, setOpen] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const qs = range === "date" && date ? `?date=${date}` : `?range=${range === "date" ? "week" : range}`;
      const res = await fetch(`/api/siparisler${qs}`);
      const d = await res.json();
      if (res.ok) {
        setOrders(d.orders || []);
        setBlob(d.blob !== false);
      }
    } finally {
      setLoading(false);
    }
  }, [range, date]);

  useEffect(() => {
    load();
  }, [load]);

  async function patch(o: Row, body: Record<string, unknown>, optimistic: Partial<Row>) {
    const prev = orders;
    setOrders(orders.map((x) => (x.orderId === o.orderId ? { ...x, ...optimistic } : x)));
    const res = await fetch(`/api/siparisler?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!res.ok) setOrders(prev);
  }

  function updatePayment(o: Row, payment: PaymentStatus) {
    let paidAmount: number | undefined;
    if (payment === "kismi") {
      const girilen = prompt(`Tahsil edilen tutar (toplam ₺${fmt(o.total)}):`, String(o.paidAmount || ""));
      if (girilen === null) return;
      paidAmount = Number(girilen.replace(",", ".")) || 0;
    }
    patch(
      o,
      paidAmount !== undefined ? { paidAmount } : { payment },
      { payment, paidAmount: payment === "odendi" ? o.total : payment === "bekliyor" ? 0 : paidAmount ?? 0 }
    );
  }

  function sendStatusWhatsApp(o: Row) {
    const phone = waDigits(o.customerPhone);
    if (!phone) return;
    const durum: Record<OrderStatus, string> = {
      Beklemede: "siparişiniz alındı, sıraya girdi.",
      "Hazırlanıyor": "siparişiniz hazırlanıyor.",
      "Hazır": "siparişiniz hazır, teslim alabilirsiniz. 🎉",
      "Teslim Edildi": "siparişiniz teslim edildi. Bizi tercih ettiğiniz için teşekkürler.",
      "İptal": "siparişiniz iptal edildi.",
    };
    const text = [
      `*${dealerName} — Sipariş ${o.orderId}*`,
      `Sayın ${o.customerName}, ${durum[o.status]}`,
      o.deliveryDate && o.status !== "Teslim Edildi" ? `Tahmini teslim: ${o.deliveryDate}` : "",
      `Toplam: ${fmt(o.total)} TL`,
      o.trackUrl ? `Sipariş takibi: ${o.trackUrl}` : "",
    ]
      .filter(Boolean)
      .join("\n");
    window.open(`https://wa.me/${phone}?text=${encodeURIComponent(text)}`, "_blank");
  }

  const q = query.trim();
  const rakamlar = q.replace(/\D/g, "");
  const filtered = q
    ? orders.filter(
        (o) =>
          eslesir(q, o.orderId, o.customerName) ||
          (rakamlar.length >= 3 && o.customerPhone.replace(/\D/g, "").includes(rakamlar))
      )
    : orders;

  const toplam = filtered.reduce((s, o) => s + (o.status === "İptal" ? 0 : o.total), 0);
  const bekleyen = filtered.reduce(
    (s, o) => s + (o.status === "İptal" ? 0 : Math.max(0, o.total - (o.paidAmount || 0))),
    0
  );

  return (
    <div>
      <div className="card" style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", padding: 14 }}>
        {(
          [
            ["today", "Bugün"],
            ["week", "Son 7 Gün"],
            ["month", "Son 30 Gün"],
            ["all", "Tümü"],
          ] as const
        ).map(([k, label]) => (
          <button key={k} className={`btn small ${range === k ? "" : "secondary"}`} onClick={() => setRange(k)}>
            {label}
          </button>
        ))}
        <input
          type="date"
          style={{ width: 160 }}
          value={date}
          onChange={(e) => {
            setDate(e.target.value);
            if (e.target.value) setRange("date");
          }}
        />
        <input
          placeholder="Ara: sipariş no / ad / telefon"
          style={{ flex: 1, minWidth: 200 }}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
      </div>

      {!blob && (
        <div className="notice info">Kalıcı depolama yapılandırılmadığı için kayıtlı sipariş listelenemiyor.</div>
      )}

      {!loading && filtered.length > 0 && (
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", fontSize: 13.5, color: "var(--text-2)", margin: "4px 0 12px" }}>
          <span>{filtered.length} sipariş</span>
          <span>Ciro: <strong>₺{fmt(toplam)}</strong></span>
          <span>Bekleyen tahsilat: <strong style={{ color: bekleyen > 0 ? "var(--error)" : "inherit" }}>₺{fmt(bekleyen)}</strong></span>
        </div>
      )}

      {loading ? (
        <p style={{ color: "var(--muted)" }}>Yükleniyor...</p>
      ) : filtered.length === 0 ? (
        <div className="card" style={{ textAlign: "center", color: "var(--muted)" }}>Bu aralıkta sipariş yok.</div>
      ) : (
        filtered.map((o) => (
          <div className="card" key={o.orderId} style={{ padding: 16, marginBottom: 12 }}>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
              <strong style={{ color: "var(--brand)" }}>{o.orderId}</strong>
              <span>{o.customerName} · {o.customerPhone}</span>
              <span style={{ color: "var(--muted)", fontSize: 13 }}>
                {new Date(o.createdAt).toLocaleString("tr-TR", { dateStyle: "short", timeStyle: "short" })}
                {o.deliveryDate && ` → Teslim: ${o.deliveryDate}`}
              </span>
              <span style={{ flex: 1 }} />
              <strong>₺{fmt(o.total)}</strong>
              <select
                className={`pay-select ${o.payment || "bekliyor"}`}
                style={{ width: 140, fontWeight: 600 }}
                value={o.payment || "bekliyor"}
                onChange={(e) => updatePayment(o, e.target.value as PaymentStatus)}
              >
                {Object.entries(PAYMENT_LABELS).map(([k, v]) => (
                  <option key={k} value={k}>{v}</option>
                ))}
              </select>
              <select
                style={{ width: 150, fontWeight: 600, color: STATUS_COLORS[o.status] || "var(--text)" }}
                value={o.status}
                onChange={(e) => patch(o, { status: e.target.value }, { status: e.target.value as OrderStatus })}
              >
                {ORDER_STATUSES.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
              <button className="btn small wa" onClick={() => sendStatusWhatsApp(o)} title="Müşteriye durum mesajı gönder">
                📲
              </button>
              <a className="btn small secondary" href={`/api/siparisler/pdf?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`} title="Üretim PDF indir">
                ⬇ PDF
              </a>
              <a className="btn small secondary" href={`/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`} title="Fişi görüntüle / yazdır">
                🖨️ Fiş
              </a>
              <button className="btn small secondary" onClick={() => setOpen(open === o.orderId ? null : o.orderId)}>
                {open === o.orderId ? "Kapat" : "Detay"}
              </button>
            </div>

            {open === o.orderId && (
              <div style={{ marginTop: 12, borderTop: "1px solid var(--border)", paddingTop: 12 }}>
                {o.items.map((it, i) => (
                  <div key={i} style={{ fontSize: 13.5, marginBottom: 8 }}>
                    <strong>{i + 1}.</strong> {it.artWidth}{it.artWidthUnit} × {it.artHeight}{it.artHeightUnit}
                    {" · Çerçeve: "}{it.frameCode}
                    {it.matCode !== "-" && (
                      <>
                        {" · Paspartu: "}{it.matType} {it.matColor !== "-" && it.matColor}
                        {it.doubleMat && ` + İç: ${it.innerMatType} ${it.innerMatColor} (montaj ${it.altMontaj}mm)`}
                        {it.zeminEnabled && ` | Zemin: ${it.zeminType} ${it.zeminColor}`}
                        {` · Kenar: ${it.matTop}/${it.matRight}/${it.matBottom}/${it.matLeft}mm`}
                      </>
                    )}
                    {it.glassType !== "Cam Yok" && ` · Cam: ${it.glassType}`}
                    {it.printType !== "Baskı Yok" && ` · Baskı: ${it.printType}`}
                    {" — "}<strong>₺{fmt(it.itemTotal)}</strong>
                  </div>
                ))}
                <div style={{ fontSize: 13.5, color: "var(--text-2)" }}>
                  {o.discount > 0 && <>İndirim: -₺{fmt(o.discount)} · </>}
                  {o.customerEmail && <>{o.customerEmail} · </>}
                  {o.customerAddress && <>{o.customerAddress} · </>}
                  {o.notes && <>Not: {o.notes} · </>}
                  {o.trackUrl && (
                    <a href={o.trackUrl} target="_blank" rel="noreferrer">Müşteri takip linki</a>
                  )}
                </div>
              </div>
            )}
          </div>
        ))
      )}
    </div>
  );
}
