import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import OrdersList from "@/components/OrdersList";
import { finansAktif } from "@/data/users";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export default async function OrdersPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/siparisler");
  if (user.role !== "staff") redirect("/portal");

  return (
    // Sipariş tablosu 9 kolon — 1200px'lik standart genişlikte İşlemler
    // sütunu taşıyordu; geniş ekranlarda bütün tabloya yer açılır.
    <main className="container" style={{ maxWidth: 1480 }}>
      <PageHeader
        icon="list"
        title="Siparişler"
        subtitle="Alınan siparişleri takip edin: durum değiştirin, yazdırın, düzenleyin."
        actions={
          <Link href="/panel" className="btn">
            <Icon name="plus" size={16} /> Yeni Sipariş
          </Link>
        }
      />
      <OrdersList eldenSatis={finansAktif()} />
    </main>
  );
}
