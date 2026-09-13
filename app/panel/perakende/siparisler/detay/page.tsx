/* eslint-disable @next/next/no-img-element */
import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { getRetailOrder } from "@/lib/retail-orders";
import PrintButton from "@/components/PrintButton";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

// Durum rozeti rengi (fişte de basılır — açık zemin, koyu yazı olarak çıkar)
const durumRozet = (s: string) =>
  s === "Teslim Edildi" ? "ok" : s === "İptal" ? "err" : s === "Hazırlanıyor" ? "warn" : "info";

export const dynamic = "force-dynamic";

export default async function RetailOrderDetailPage({
  searchParams,
}: {
  searchParams: { d?: string; id?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/perakende/siparisler");
  if (user.role !== "staff") redirect("/portal");

  const dateKey = searchParams.d || "";
  const orderId = searchParams.id || "";
  const order = dateKey && orderId ? await getRetailOrder(dateKey, orderId) : null;

  const geriLink = (
    <Link href="/panel/perakende/siparisler" className="btn secondary">
      <Icon name="chevron-left" size={16} /> Perakende Siparişler
    </Link>
  );

  if (!order) {
    return (
      <main className="container">
        <PageHeader icon="file-text" title="Perakende Sipariş Fişi" actions={geriLink} />
        <div className="notice err">Sipariş bulunamadı.</div>
      </main>
    );
  }

  const pdfHref = `/api/perakende/orders/pdf?d=${order.dateKey}&id=${encodeURIComponent(order.orderId)}`;

  return (
    <main className="container" style={{ maxWidth: 820 }}>
      <PageHeader
        icon="file-text"
        title={<>Perakende Sipariş Fişi — {order.orderId}</>}
        subtitle={
          <>Müşteri: {order.customerName} · Personel: {order.employee}</>
        }
        actions={
          <>
            {geriLink}
            <a href={pdfHref} className="btn secondary">
              <Icon name="download" size={16} /> Üretim PDF
            </a>
            <PrintButton />
          </>
        }
      />

      <div className="card" id="print-area">
        {/* Başlık */}
        <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap", borderBottom: "3px solid var(--brand)", paddingBottom: 14, marginBottom: 18 }}>
          <img src="/logo.png" alt="Olga Çerçeve" style={{ height: 40, width: "auto" }} />
          <span style={{ flex: 1 }} />
          <div className="stack" style={{ gap: 3, alignItems: "flex-end", textAlign: "right" }}>
            <div style={{ fontWeight: 800, fontSize: 17, color: "var(--brand)" }}>
              PERAKENDE SİPARİŞ FİŞİ
            </div>
            <div style={{ fontSize: 13, color: "var(--text-2)" }}>
              {order.orderId} ·{" "}
              {new Date(order.createdAt).toLocaleString("tr-TR", { dateStyle: "medium", timeStyle: "short" })}
            </div>
            <div className="row" style={{ gap: 6, fontSize: 12.5, color: "var(--muted)" }}>
              Durum:{" "}
              <span className={`badge ${durumRozet(order.status)}`}>{order.status}</span>
            </div>
          </div>
        </div>

        {/* Müşteri */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 18 }}>
          <div style={{ background: "var(--surface-2)", border: "1px solid var(--hairline)", borderRadius: "var(--radius-xs)", padding: "12px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Müşteri</div>
            <div style={{ fontWeight: 700 }}>{order.customerName}</div>
            <div style={{ fontSize: 13 }}>{order.customerPhone}</div>
            {order.customerAddress && (
              <div style={{ fontSize: 12.5, color: "var(--text-2)" }}>{order.customerAddress}</div>
            )}
          </div>
          <div style={{ background: "var(--surface-2)", border: "1px solid var(--hairline)", borderRadius: "var(--radius-xs)", padding: "12px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Sipariş Bilgileri</div>
            <div style={{ fontSize: 13 }}>Personel: <strong>{order.employee}</strong></div>
            <div style={{ fontSize: 13 }}>Teslim: <strong>{order.deliveryDate || "-"}</strong></div>
          </div>
        </div>

        {/* Kalemler */}
        {order.items.map((it, i) => (
          <div key={i} style={{ border: "1px solid var(--border)", borderRadius: "var(--radius-sm)", padding: "12px 16px", marginBottom: 10 }}>
            <div className="row" style={{ alignItems: "baseline" }}>
              <strong style={{ color: "var(--brand)" }}>{order.items.length > 1 ? `#${i + 1}` : "Ürün"}</strong>
              <span style={{ fontWeight: 700 }}>
                {it.artWidth} {it.artWidthUnit} × {it.artHeight} {it.artHeightUnit}
              </span>
              <span>Çerçeve: <strong>{it.frameCode}</strong></span>
              <span style={{ flex: 1 }} />
              <strong className="num">₺{fmt(it.itemTotal)}</strong>
            </div>
            <div style={{ fontSize: 13, color: "var(--text-2)", marginTop: 4 }}>
              {it.kasa && (
                <>
                  <b>Kasa (kanvas) çerçeve</b> — cam ve paspartu uygulanmaz
                  <br />
                </>
              )}
              {(Number(it.pencereSayisi) || 1) > 1 && (
                <>
                  <b>{it.pencereSayisi} pencere</b> (
                  {(it.pencereDuzen || "").replace("x", "×")}
                  {it.pencereAralik ? `, aralık ${it.pencereAralik} mm` : ""}) — ölçü
                  tek fotoğrafındır
                  <br />
                </>
              )}
              {it.matType !== "Paspartu Yok" ? (
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

        {/* Toplamlar */}
        <div className="num" style={{ marginTop: 16, marginLeft: "auto", maxWidth: 320 }}>
          <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 14 }}>
            <span>Ara Toplam</span>
            <span>₺{fmt(order.gross)}</span>
          </div>
          {order.discount > 0 && (
            <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 14, color: "var(--error)" }}>
              <span>İndirim</span>
              <span>-₺{fmt(order.discount)}</span>
            </div>
          )}
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              padding: "10px 14px",
              marginTop: 6,
              borderRadius: "var(--radius-xs)",
              background: "var(--brand-btn)",
              color: "var(--on-brand)",
              fontWeight: 800,
              fontSize: 16,
            }}
          >
            <span>GENEL TOPLAM</span>
            <span>₺{fmt(order.total)}</span>
          </div>
          {/* Kapora alındıysa fişte kalan bakiye açıkça görünür */}
          {(Number(order.paidAmount) || 0) > 0 && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0 2px", fontSize: 14 }}>
                <span>Kapora / Ödenen</span>
                <span>-₺{fmt(Number(order.paidAmount) || 0)}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "2px 0", fontSize: 15, fontWeight: 800 }}>
                <span>Teslimde Kalan</span>
                <span>
                  ₺{fmt(Math.max(0, order.total - (Number(order.paidAmount) || 0)))}
                </span>
              </div>
            </>
          )}
        </div>

        {order.notes && (
          <div className="notice info" style={{ marginTop: 16 }}>
            <strong>NOT:</strong> {order.notes}
          </div>
        )}

        <div style={{ marginTop: 22, paddingTop: 12, borderTop: "1px solid var(--border)", display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: "4px 16px", fontSize: 12, color: "var(--muted)" }}>
          <span>OLGA Çerçeve</span>
          <span>0850 305 75 45</span>
          <span>www.olgacerceve.com</span>
        </div>
      </div>
    </main>
  );
}
