import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import RetailOrdersList from "@/components/RetailOrdersList";
import { isOwner } from "@/data/users";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export default async function PerakendeSiparislerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/perakende/siparisler");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1100 }}>
      <PageHeader
        icon="image"
        title="Perakende Siparişler"
        subtitle="Çerçeveletme siparişlerini görüntüleyin ve durumlarını güncelleyin."
        actions={
          <Link href="/panel/perakende" className="btn">
            <Icon name="plus" size={16} /> Yeni Perakende Sipariş
          </Link>
        }
      />
      <RetailOrdersList patronGonderim={isOwner(user.username)} />
    </main>
  );
}
