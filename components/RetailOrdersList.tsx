"use client";

// Perakende sipariş listesi — filtre, arama ve durum güncelleme.

import { useCallback, useEffect, useState } from "react";
import { RETAIL_STATUSES, type RetailStatus } from "@/data/perakende";
import { PAYMENT_LABELS, type PaymentStatus } from "@/lib/orders";
import type { SavedRetailOrder } from "@/lib/retail-orders";
import { eslesir } from "@/lib/search-norm";
import Icon from "@/components/shell/Icon";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

// Durum seçicisinin yazı rengi — tema tokenları (koyu/açık temada okunur)
const STATUS_COLORS: Record<RetailStatus, string> = {
  Beklemede: "var(--warning)",
  "Hazırlanıyor": "var(--info)",
  "Hazır": "var(--success)",
  "Teslim Edildi": "var(--muted)",
  "İptal": "var(--error)",
};

export default function RetailOrdersList({
  // Patrona WhatsApp ile fiş gönderme düğmesi — yalnızca sahipler (sayfa geçirir)
  patronGonderim = false,
}: {
  patronGonderim?: boolean;
} = {}) {
  const [range, setRange] = useState<"today" | "week" | "date">("today");
  const [date, setDate] = useState("");
  const [query, setQuery] = useState("");
  const [orders, setOrders] = useState<SavedRetailOrder[]>([]);
  const [loading, setLoading] = useState(true);
  const [blob, setBlob] = useState(true);
  const [open, setOpen] = useState<string | null>(null);
  // Patrona WhatsApp ile fiş gönderimi — sipariş başına durum metni
  const [waDurum, setWaDurum] = useState<Record<string, string>>({});

  async function patronaGonder(o: SavedRetailOrder) {
    if (!window.confirm(`${o.orderId} (${o.customerName}) fişi patrona WhatsApp ile PDF olarak gönderilecek. Devam edilsin mi?`)) return;
    setWaDurum((d) => ({ ...d, [o.orderId]: "Gönderiliyor…" }));
    try {
      const res = await fetch("/api/whatsapp/patron", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ orders: [{ d: o.dateKey, id: o.orderId }] }),
      });
      const j = await res.json().catch(() => null);
      const r = j?.sonuclar?.[0];
      setWaDurum((d) => ({ ...d, [o.orderId]: r?.ok ? "Patrona gönderildi ✓" : `Gönderilemedi: ${r?.hata || j?.error || "bilinmeyen hata"}` }));
    } catch {
      setWaDurum((d) => ({ ...d, [o.orderId]: "Gönderilemedi: sunucuya ulaşılamadı" }));
    }
  }

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const qs =
        range === "date" && date ? `?date=${date}` : `?range=${range === "date" ? "today" : range}`;
      const res = await fetch(`/api/perakende/orders${qs}`);
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

  async function updateStatus(o: SavedRetailOrder, status: RetailStatus) {
    const prev = orders;
    setOrders(orders.map((x) => (x.orderId === o.orderId ? { ...x, status } : x)));
    const res = await fetch(
      `/api/perakende/orders?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ status }),
      }
    );
    if (!res.ok) setOrders(prev);
  }

  async function updatePayment(o: SavedRetailOrder, payment: PaymentStatus) {
    let paidAmount: number | undefined;
    if (payment === "kismi") {
      const girilen = prompt(
        `Tahsil edilen tutar (toplam ₺${fmt(o.total)}):`,
        String(o.paidAmount || "")
      );
      if (girilen === null) return;
      paidAmount = Number(girilen.replace(",", ".")) || 0;
    }
    const prev = orders;
    setOrders(
      orders.map((x) =>
        x.orderId === o.orderId
          ? {
              ...x,
              payment,
              paidAmount:
                payment === "odendi" ? o.total : payment === "bekliyor" ? 0 : paidAmount ?? 0,
            }
          : x
      )
    );
    const res = await fetch(
      `/api/perakende/orders?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(paidAmount !== undefined ? { paidAmount } : { payment }),
      }
    );
    if (!res.ok) setOrders(prev);
  }

  // Türkçe karakter ve büyük/küçük harf duyarsız arama (bkz. lib/search-norm)
  const q = query.trim();
  const rakamlar = q.replace(/\D/g, "");
  const filtered = q
    ? orders.filter(
        (o) =>
          eslesir(q, o.orderId, o.customerName) ||
          (rakamlar.length >= 3 &&
            o.customerPhone.replace(/\D/g, "").includes(rakamlar))
      )
    : orders;

  return (
    <div>
      <div className="card pad-sm row">
        {/* Tarih aralığı — segmentli kontrol (.seg); aktif seçenek .active */}
        <div className="seg" role="group" aria-label="Tarih aralığı">
          <button type="button" className={range === "today" ? "active" : ""} onClick={() => setRange("today")}>
            Bugün
          </button>
          <button type="button" className={range === "week" ? "active" : ""} onClick={() => setRange("week")}>
            Son 7 Gün
          </button>
        </div>
        <input
          type="date"
          style={{ flex: "1 1 160px", minWidth: 160, maxWidth: "100%" }}
          value={date}
          onChange={(e) => {
            setDate(e.target.value);
            if (e.target.value) setRange("date");
          }}
        />
        <input
          placeholder="Ara: sipariş no / ad / telefon"
          style={{ flex: "3 1 200px", minWidth: 200 }}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
      </div>

      {!blob && (
        <div className="notice info">
          Kalıcı depolama yapılandırılmadığı için kayıtlı sipariş listelenemiyor.
        </div>
      )}

      {loading ? (
        <p className="muted">Yükleniyor...</p>
      ) : filtered.length === 0 ? (
        <div className="card">
          <div className="empty">
            <div className="empty-icon">
              <Icon name="inbox" size={22} />
            </div>
            Bu aralıkta perakende sipariş yok.
          </div>
        </div>
      ) : (
        filtered.map((o) => (
          <div className="card pad-sm" key={o.orderId}>
            <div className="row">
              <strong style={{ color: "var(--brand)" }}>{o.orderId}</strong>
              <span style={{ minWidth: 0, overflowWrap: "anywhere" }}>{o.customerName} · {o.customerPhone}</span>
              <span style={{ color: "var(--muted)", fontSize: 13 }}>
                {new Date(o.createdAt).toLocaleString("tr-TR", { dateStyle: "short", timeStyle: "short" })}
                {o.deliveryDate && ` → Teslim: ${o.deliveryDate}`}
              </span>
              <span className="spacer" />
              <strong className="num">₺{fmt(o.total)}</strong>
              <select
                className={`pay-select ${o.payment || "bekliyor"}`}
                style={{ width: 140, maxWidth: "100%", fontWeight: 600 }}
                value={o.payment || "bekliyor"}
                onChange={(e) => updatePayment(o, e.target.value as PaymentStatus)}
              >
                {Object.entries(PAYMENT_LABELS).map(([k, v]) => (
                  <option key={k} value={k}>{v}</option>
                ))}
              </select>
              <select
                style={{
                  width: 150,
                  maxWidth: "100%",
                  fontWeight: 600,
                  color: STATUS_COLORS[o.status] || "var(--text)",
                }}
                value={o.status}
                onChange={(e) => updateStatus(o, e.target.value as RetailStatus)}
              >
                {RETAIL_STATUSES.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
              <a
                className="btn small secondary"
                href={`/api/perakende/orders/pdf?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                title="Üretim PDF indir"
              >
                <Icon name="download" size={14} /> PDF
              </a>
              {patronGonderim && (
                <button
                  type="button"
                  className="btn small secondary"
                  title="Fişi patrona WhatsApp ile PDF olarak gönder"
                  disabled={waDurum[o.orderId] === "Gönderiliyor…"}
                  onClick={() => patronaGonder(o)}
                >
                  <Icon name="message" size={14} /> Patrona
                </button>
              )}
              {patronGonderim && waDurum[o.orderId] && (
                <span className="text-2" style={{ fontSize: 12.5, color: waDurum[o.orderId].startsWith("Gönderilemedi") ? "var(--error)" : "var(--success)" }}>
                  {waDurum[o.orderId]}
                </span>
              )}
              <a
                className="btn small secondary"
                href={`/panel/perakende/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                title="Fişi görüntüle / yazdır"
              >
                <Icon name="printer" size={14} /> Fiş
              </a>
              <button
                type="button"
                className="btn small secondary"
                onClick={() => setOpen(open === o.orderId ? null : o.orderId)}
              >
                <Icon name={open === o.orderId ? "chevron-down" : "chevron-right"} size={14} />
                {open === o.orderId ? "Kapat" : "Detay"}
              </button>
            </div>

            {open === o.orderId && (
              <div style={{ marginTop: 12, borderTop: "1px solid var(--border)", paddingTop: 12 }}>
                {o.items.map((it, i) => (
                  <div key={i} style={{ fontSize: 13.5, marginBottom: 8 }}>
                    <strong>{i + 1}.</strong> {it.artWidth}{it.artWidthUnit} × {it.artHeight}{it.artHeightUnit}
                    {" · Çerçeve: "}{it.frameCode}
                    {it.matType !== "Paspartu Yok" && (
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
                  Personel: {o.employee}
                  {o.customerEmail && <> · {o.customerEmail}</>}
                  {o.notes && <> · Not: {o.notes}</>}
                </div>
              </div>
            )}
          </div>
        ))
      )}
    </div>
  );
}
