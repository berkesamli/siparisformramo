// Sipariş fişi — bayi markalı; hem bayinin fiş sayfasında hem müşteri takip
// sayfasında kullanılır (sunucu bileşeni).
import type { SavedOrder } from "@/lib/orders";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export interface ReceiptDealer {
  name: string;
  phone: string;
  website?: string;
  address?: string;
}

export default function OrderReceipt({
  order,
  dealer,
  customerView = false,
}: {
  order: SavedOrder;
  dealer: ReceiptDealer;
  customerView?: boolean;
}) {
  return (
    <div className="card" style={{ padding: 30 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 16, borderBottom: "3px solid var(--brand)", paddingBottom: 14, marginBottom: 18 }}>
        <div>
          <div style={{ fontWeight: 800, fontSize: 20, color: "var(--brand)" }}>{dealer.name}</div>
          <div style={{ fontSize: 12.5, color: "var(--muted)" }}>
            {dealer.phone}{dealer.website && ` · ${dealer.website}`}
          </div>
        </div>
        <span style={{ flex: 1 }} />
        <div style={{ textAlign: "right" }}>
          <div style={{ fontWeight: 800, fontSize: 17, color: "var(--brand)" }}>SİPARİŞ FİŞİ</div>
          <div style={{ fontSize: 13, color: "var(--text-2)" }}>
            {order.orderId} · {new Date(order.createdAt).toLocaleString("tr-TR", { dateStyle: "medium", timeStyle: "short" })}
          </div>
          <div style={{ fontSize: 12.5, color: "var(--muted)" }}>Durum: <strong>{order.status}</strong></div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 18 }}>
        <div style={{ background: "var(--input)", borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Müşteri</div>
          <div style={{ fontWeight: 700 }}>{order.customerName}</div>
          <div style={{ fontSize: 13 }}>{order.customerPhone}</div>
          {order.customerAddress && <div style={{ fontSize: 12.5, color: "var(--text-2)" }}>{order.customerAddress}</div>}
        </div>
        <div style={{ background: "var(--input)", borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Sipariş Bilgileri</div>
          <div style={{ fontSize: 13 }}>Teslim: <strong>{order.deliveryDate || "-"}</strong></div>
          {!customerView && <div style={{ fontSize: 13 }}>Kaydeden: <strong>{order.createdBy}</strong></div>}
          {order.paidAmount > 0 && (
            <div style={{ fontSize: 13 }}>Tahsil edilen: <strong>₺{fmt(order.paidAmount)}</strong></div>
          )}
        </div>
      </div>

      {order.items.map((it, i) => (
        <div key={i} style={{ border: "1px solid var(--border)", borderRadius: 12, padding: "12px 16px", marginBottom: 10 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "baseline" }}>
            <strong style={{ color: "var(--brand)" }}>{order.items.length > 1 ? `#${i + 1}` : "Ürün"}</strong>
            <span style={{ fontWeight: 700 }}>{it.artWidth} {it.artWidthUnit} × {it.artHeight} {it.artHeightUnit}</span>
            <span>Çerçeve: <strong>{it.frameCode}</strong></span>
            <span style={{ flex: 1 }} />
            <strong>₺{fmt(it.itemTotal)}</strong>
          </div>
          <div style={{ fontSize: 13, color: "var(--text-2)", marginTop: 4 }}>
            {it.matCode !== "-" ? (
              <>
                Paspartu: {it.matType}
                {it.matColor !== "-" && ` (${it.matColor})`}
                {it.doubleMat && ` + İç: ${it.innerMatType} (${it.innerMatColor}) · montaj ${it.altMontaj}mm`}
                {it.zeminEnabled && ` | Zemin: ${it.zeminType} (${it.zeminColor})`}
                {` · Kenarlar: Ü${it.matTop}/S${it.matRight}/A${it.matBottom}/S${it.matLeft}mm`}
                <br />
              </>
            ) : (
              <>Paspartu: Yok<br /></>
            )}
            Cam: {it.glassType}
            {it.printType !== "Baskı Yok" && <> · Baskı: {it.printType}</>}
          </div>
        </div>
      ))}

      <div style={{ marginTop: 16, marginLeft: "auto", maxWidth: 320 }}>
        <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 14 }}>
          <span>Ara Toplam</span><span>₺{fmt(order.gross)}</span>
        </div>
        {order.discount > 0 && (
          <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 14, color: "var(--error)" }}>
            <span>İndirim</span><span>-₺{fmt(order.discount)}</span>
          </div>
        )}
        <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 14px", marginTop: 6, borderRadius: 10, background: "linear-gradient(135deg, var(--brand), var(--brand-dark))", color: "#fff", fontWeight: 800, fontSize: 16 }}>
          <span>GENEL TOPLAM</span><span>₺{fmt(order.total)}</span>
        </div>
      </div>

      {order.notes && (
        <div className="notice info" style={{ marginTop: 16 }}>
          <strong>NOT:</strong> {order.notes}
        </div>
      )}

      <div style={{ marginTop: 22, paddingTop: 12, borderTop: "1px solid var(--border)", display: "flex", justifyContent: "space-between", fontSize: 12, color: "var(--muted)", flexWrap: "wrap", gap: 8 }}>
        <span>{dealer.name}</span>
        <span>{dealer.phone}</span>
        <span>{dealer.website || "Olga Çerçeve yetkili bayisi"}</span>
      </div>
    </div>
  );
}
