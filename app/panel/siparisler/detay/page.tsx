import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { getOrder, STATUS_LABELS } from "@/lib/orders";
import PrintButton from "@/components/PrintButton";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

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

  if (!order) {
    return (
      <main className="container">
        <div className="notice err">Sipariş bulunamadı.</div>
        <Link href="/panel/siparisler" className="btn small secondary">
          ← Siparişler
        </Link>
      </main>
    );
  }

  return (
    <main className="container" style={{ maxWidth: 820 }}>
      <div className="no-print" style={{ display: "flex", gap: 10, marginBottom: 14 }}>
        <Link href="/panel/siparisler" className="btn small secondary">
          ← Siparişler
        </Link>
        <span style={{ flex: 1 }} />
        <Link
          href={`/panel/siparisler/duzenle?d=${order.dateKey}&id=${encodeURIComponent(order.orderId)}`}
          className="btn small secondary"
        >
          ✏️ Düzenle
        </Link>
        <a
          className="btn small"
          href={`/api/orders/pdf?d=${order.dateKey}&id=${encodeURIComponent(order.orderId)}`}
        >
          ⬇ PDF İndir
        </a>
        <PrintButton />
      </div>

      <div className="card" id="print-area">
        <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
          <div>
            <h1 style={{ fontSize: 22, color: "var(--brand-light)" }}>OLGA ÇERÇEVE</h1>
            <p style={{ color: "var(--text-2)" }}>Sipariş Fişi</p>
          </div>
          <div style={{ textAlign: "right", fontSize: 13.5 }}>
            <div><strong>Sipariş No:</strong> {order.orderId}</div>
            <div>
              <strong>Tarih:</strong>{" "}
              {new Date(order.createdAt).toLocaleString("tr-TR", {
                dateStyle: "medium",
                timeStyle: "short",
                timeZone: "Europe/Istanbul",
              })}
            </div>
            <div><strong>Durum:</strong> {STATUS_LABELS[order.status]}</div>
          </div>
        </div>

        <hr style={{ border: "none", borderTop: "1px solid var(--border)", margin: "14px 0" }} />

        <div style={{ display: "flex", gap: 30, flexWrap: "wrap", fontSize: 14 }}>
          <div><strong>Müşteri:</strong> {order.customer || "—"}</div>
          <div><strong>Çalışan:</strong> {order.employee}</div>
          {order.note && <div><strong>Not:</strong> {order.note}</div>}
        </div>

        <table style={{ marginTop: 16 }}>
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
                <td>₺ {fmt(l.unitPriceTL)}</td>
                <td>₺ {fmt(l.lineTotal)}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <table style={{ marginTop: 14, maxWidth: 340, marginLeft: "auto" }}>
          <tbody>
            <tr>
              <td>Ara Toplam</td>
              <td style={{ textAlign: "right" }}>₺ {fmt(order.gross)}</td>
            </tr>
            <tr>
              <td>İskonto (%{order.discountPct})</td>
              <td style={{ textAlign: "right" }}>₺ {fmt(order.discount)}</td>
            </tr>
            <tr>
              <td>KDV</td>
              <td style={{ textAlign: "right" }}>
                {order.vatApplied ? `%20 — ₺ ${fmt(order.vatAmount)}` : "Uygulanmadı"}
              </td>
            </tr>
            <tr>
              <td><strong>GENEL TOPLAM</strong></td>
              <td style={{ textAlign: "right" }}>
                <strong>₺ {fmt(order.net)}</strong>
              </td>
            </tr>
          </tbody>
        </table>

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
