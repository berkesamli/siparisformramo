import Link from "next/link";
import { redirect } from "next/navigation";
import { getDealerSession } from "@/lib/auth";
import OrdersList from "@/components/OrdersList";

export const dynamic = "force-dynamic";

export default async function SiparislerPage() {
  const s = await getDealerSession();
  if (!s) redirect("/giris?next=/panel/siparisler");
  return (
    <main className="container" style={{ maxWidth: 1100 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ marginBottom: 4 }}>Siparişler</h1>
        <span style={{ flex: 1 }} />
        <Link href="/panel/cerceve" className="btn small">➕ Yeni Sipariş</Link>
      </div>
      <p className="subtitle">Siparişlerinizi görüntüleyin, durum ve ödeme bilgisini güncelleyin, müşteriye WhatsApp'tan bildirin.</p>
      <OrdersList dealerName={s.dealer.name} />
    </main>
  );
}
