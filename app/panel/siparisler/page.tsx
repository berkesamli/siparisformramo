import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import OrdersList from "@/components/OrdersList";
import { finansAktif } from "@/data/users";

export default async function OrdersPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/siparisler");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ flex: 1 }}>
          <h1>Siparişler</h1>
          <p className="subtitle">
            Alınan siparişleri takip edin: durum değiştirin, yazdırın, düzenleyin.
          </p>
        </div>
        <Link href="/panel" className="btn small secondary">
          + Yeni Sipariş
        </Link>
      </div>
      <OrdersList eldenSatis={finansAktif()} />
    </main>
  );
}
