import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import OrdersList from "@/components/OrdersList";
import { isOwner } from "@/data/users";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function TamamlananSiparislerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/siparisler/tamamlanan");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1480 }}>
      <PageHeader
        icon="check-circle"
        title="Tamamlanan Siparişler"
        subtitle={
          <>
            Durumu “Tamamlandı”, ödemesi alınmış ve kontrol edilmiş siparişler —
            aktif listeden çıkıp buraya düşer.
          </>
        }
        actions={
          <Link href="/panel/siparisler" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Aktif Siparişler
          </Link>
        }
      />
      <OrdersList tamamlananlar patronGonderim={isOwner(user.username)} />
    </main>
  );
}
