import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { getOrder, STATUS_LABELS, type OrderStatus } from "@/lib/orders";
import PrintButton from "@/components/PrintButton";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

// Durum rozeti rengi (fişte de basılır — açık zemin, koyu yazı olarak çıkar)
const durumRozet = (s: OrderStatus) =>
  s === "tamamlandi" ? "ok" : s === "iptal" ? "err" : s === "hazirlaniyor" ? "warn" : s === "yarim" ? "yarim" : "info";

export const dynamic = "force-dynamic";

export default async function OrderDetailPage({
  searchParams,
}: {
  searchParams: { d?: string; id?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/siparisler");
  if (user.role !== "staff") redirect("/portal");

  const dateKey = searchParams.d || "";
  const orderId = searchParams.id || "";
  const order = dateKey && orderId ? await getOrder(dateKey, orderId) : null;

  const geriLink = (
    // ?donus=1: liste kaldığı filtre ve kaydırma konumuyla açılır (OrdersList hatırlama)
    <Link href="/panel/siparisler?donus=1" className="btn secondary">
      <Icon name="chevron-left" size={16} /> Siparişler
    </Link>
  );

  if (!order) {
    return (
      <main className="container">
        <PageHeader icon="file-text" title="Sipariş Fişi" actions={geriLink} />
        <div className="notice err">Sipariş bulunamadı.</div>
      </main>
    );
  }

  return (
    <main className="container" style={{ maxWidth: 820 }}>
      <PageHeader
        icon="file-text"
        title={<>Sipariş Fişi — {order.orderId}</>}
        subtitle={
          <>Müşteri: {order.customer || "—"} · Oluşturan: {order.employee}</>
        }
        actions={
          <>
            {geriLink}
            <Link
              href={`/panel/siparisler/duzenle?d=${order.dateKey}&id=${encodeURIComponent(order.orderId)}`}
              className="btn secondary"
            >
              <Icon name="edit" size={16} /> Düzenle
            </Link>
            <a
              className="btn"
              href={`/api/orders/pdf?d=${order.dateKey}&id=${encodeURIComponent(order.orderId)}`}
            >
              <Icon name="download" size={16} /> PDF İndir
            </a>
            <PrintButton />
          </>
        }
      />

      <div className="card" id="print-area">
        <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
          <div>
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img
              src="/logo.png"
              alt="Olga Çerçeve"
              style={{ height: 44, width: "auto", display: "block" }}
            />
            <p style={{ color: "var(--text-2)", marginTop: 6 }}>Sipariş Fişi</p>
          </div>
          <div className="stack" style={{ gap: 4, alignItems: "flex-end", textAlign: "right", fontSize: 13.5 }}>
            <div><strong>Sipariş No:</strong> {order.orderId}</div>
            <div>
              <strong>Tarih:</strong>{" "}
              {new Date(order.createdAt).toLocaleString("tr-TR", {
                dateStyle: "medium",
                timeStyle: "short",
                timeZone: "Europe/Istanbul",
              })}
            </div>
            <div className="row" style={{ gap: 6 }}>
              <strong>Durum:</strong>{" "}
              <span className={`badge ${durumRozet(order.status)}`}>
                {STATUS_LABELS[order.status]}
              </span>
            </div>
          </div>
        </div>

        <hr style={{ border: "none", borderTop: "1px solid var(--border)", margin: "14px 0" }} />

        <div className="row" style={{ gap: "8px 30px", fontSize: 14 }}>
          <div><strong>Müşteri:</strong> {order.customer || "—"}</div>
          <div><strong>Çalışan:</strong> {order.employee}</div>
          {order.rate > 0 && (
            <div><strong>Dolar Kuru:</strong> ₺ {fmt(order.rate)}</div>
          )}
          {order.euroRate > 0 && (
            <div><strong>Euro Kuru:</strong> ₺ {fmt(order.euroRate)}</div>
          )}
        </div>

        <div className="table-wrap" style={{ marginTop: 16 }}>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Ürün</th>
                <th>Birim</th>
                <th>Birim Fiyat</th>
                <th>Tutar</th>
              </tr>
            </thead>
            <tbody>
              {order.lines.map((l, i) => (
                <tr key={i}>
                  <td>{i + 1}</td>
                  <td>{l.name}</td>
                  <td>{l.unitText}</td>
                  <td className="num">₺ {fmt(l.unitPriceTL)}</td>
                  <td className="num">₺ {fmt(l.lineTotal)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <table style={{ marginTop: 14, maxWidth: 340, marginLeft: "auto" }}>
          <tbody>
            <tr>
              <td>Ara Toplam</td>
              <td className="num" style={{ textAlign: "right" }}>₺ {fmt(order.gross)}</td>
            </tr>
            <tr>
              <td>İskonto (%{order.discountPct})</td>
              <td className="num" style={{ textAlign: "right" }}>₺ {fmt(order.discount)}</td>
            </tr>
            <tr>
              <td>KDV</td>
              <td className="num" style={{ textAlign: "right" }}>
                {order.vatApplied ? `%20 — ₺ ${fmt(order.vatAmount)}` : "Uygulanmadı"}
              </td>
            </tr>
            <tr>
              <td><strong>GENEL TOPLAM</strong></td>
              <td className="num" style={{ textAlign: "right" }}>
                <strong>₺ {fmt(order.net)}</strong>
              </td>
            </tr>
          </tbody>
        </table>

        {order.note && (
          <div
            style={{
              marginTop: 16,
              padding: "12px 16px",
              background: "var(--surface-2)",
              border: "1px solid var(--hairline)",
              borderRadius: "var(--radius-xs)",
              fontSize: 13.5,
            }}
          >
            <span
              style={{
                color: "var(--brand)",
                fontWeight: 700,
                fontSize: 11,
                letterSpacing: "0.1em",
                display: "block",
                marginBottom: 4,
              }}
            >
              NOT
            </span>
            {order.note}
          </div>
        )}

        <p style={{ color: "var(--muted)", fontSize: 12, marginTop: 18 }}>
          Olga Çerçeve — Sipariş Hattı: 0850 305 75 45 · olgacerceve.com
        </p>
      </div>

      <p className="no-print" style={{ color: "var(--muted)", fontSize: 12.5, marginTop: 10 }}>
        💡 &quot;Yazdır / PDF&quot; butonunda hedef olarak &quot;PDF olarak
        kaydet&quot; seçerseniz dosya olarak iner.
      </p>
    </main>
  );
}
