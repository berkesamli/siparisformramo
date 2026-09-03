import Link from "next/link";
import { redirect } from "next/navigation";
import { getDealerSession } from "@/lib/auth";
import { dealerCanOrder, SUBSCRIPTION_LABELS } from "@/lib/dealers";
import { listOrders } from "@/lib/orders";
import { lastNDateKeys } from "@/lib/store";

export const dynamic = "force-dynamic";

const fmt = (n: number) => (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });

export default async function PanelPage() {
  const s = await getDealerSession();
  if (!s) redirect("/giris?next=/panel");
  const d = s.dealer;
  const can = dealerCanOrder(d);
  const orders = await listOrders(d.slug, lastNDateKeys(30));
  const aktif = orders.filter((o) => o.status !== "Teslim Edildi" && o.status !== "İptal");
  const ciro = orders.reduce((sum, o) => sum + (o.status === "İptal" ? 0 : o.total), 0);
  const bekleyen = orders.reduce(
    (sum, o) => sum + (o.status === "İptal" ? 0 : Math.max(0, o.total - (o.paidAmount || 0))),
    0
  );

  return (
    <main className="container" style={{ maxWidth: 1000 }}>
      <h1 style={{ marginBottom: 4 }}>{d.name}</h1>
      <p className="subtitle">Bayi paneli — online çerçeve fiyatlandırma ve sipariş takibi</p>

      {!can.ok && <div className="notice err">{can.reason}</div>}
      {d.subscription.status === "odeme_bekliyor" && can.ok && (
        <div className="notice info">
          Abonelik ödemeniz bekleniyor{d.subscription.paidUntil ? ` (son ödenen dönem: ${d.subscription.paidUntil})` : ""}.
          Ödeme yapılmazsa 7 gün sonra yeni sipariş kaydı kapanır.
        </div>
      )}

      <div className="rw-grid2" style={{ marginTop: 16 }}>
        <Link href="/panel/cerceve" className="card" style={{ textDecoration: "none", color: "inherit" }}>
          <div style={{ fontSize: 34 }}>🖼️</div>
          <h2 style={{ margin: "6px 0 4px" }}>Online Çerçeve</h2>
          <p style={{ margin: 0, color: "var(--text-2)", fontSize: 14 }}>
            Ölçü, çerçeve, paspartu, cam ve baskı seçin; fiyat anında çıkar. Müşteriye WhatsApp teklifi ve PDF gönderin.
          </p>
        </Link>
        <Link href="/panel/siparisler" className="card" style={{ textDecoration: "none", color: "inherit" }}>
          <div style={{ fontSize: 34 }}>📋</div>
          <h2 style={{ margin: "6px 0 4px" }}>Siparişler</h2>
          <p style={{ margin: 0, color: "var(--text-2)", fontSize: 14 }}>
            Son 30 gün: <strong>{orders.length}</strong> sipariş · aktif <strong>{aktif.length}</strong> · ciro ₺{fmt(ciro)}
            {bekleyen > 0 && <> · bekleyen tahsilat ₺{fmt(bekleyen)}</>}
          </p>
        </Link>
        <Link href="/panel/ayarlar" className="card" style={{ textDecoration: "none", color: "inherit" }}>
          <div style={{ fontSize: 34 }}>⚙️</div>
          <h2 style={{ margin: "6px 0 4px" }}>Fiyat & Ayarlar</h2>
          <p style={{ margin: 0, color: "var(--text-2)", fontSize: 14 }}>
            Çerçeve çarpanı, cam / paspartu / baskı fiyatları, kur, firma bilgileri.
          </p>
        </Link>
        <div className="card">
          <div style={{ fontSize: 34 }}>🪪</div>
          <h2 style={{ margin: "6px 0 4px" }}>Abonelik</h2>
          <p style={{ margin: 0, color: "var(--text-2)", fontSize: 14 }}>
            Durum: <strong>{SUBSCRIPTION_LABELS[d.subscription.status]}</strong>
            {d.subscription.paidUntil && <> · ödenmiş dönem: {d.subscription.paidUntil}</>}
            {d.subscription.note && <><br />{d.subscription.note}</>}
          </p>
          <p style={{ margin: "8px 0 0", color: "var(--muted)", fontSize: 12.5 }}>
            Olga Çerçeve'den yaptığınız aylık toptan alım eşiği aşınca panel ücretsizdir.
          </p>
        </div>
      </div>
    </main>
  );
}
