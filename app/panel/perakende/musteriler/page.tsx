import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import RetailCustomerManager from "@/components/RetailCustomerManager";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function PerakendeMusterilerPage({
  searchParams,
}: {
  searchParams?: { q?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/perakende/musteriler");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <PageHeader
        icon="user"
        title="Perakende Müşterileri"
        subtitle={
          <>
            Mağaza müşterileri — etiket/toptan defterinden ayrıdır. Sipariş
            kaydedilince müşteri telefon numarasıyla buraya kendiliğinden işlenir.
          </>
        }
        actions={
          <>
            <Link href="/panel/perakende/siparisler" className="btn secondary">
              <Icon name="image" size={16} /> Perakende Siparişler
            </Link>
            <Link href="/panel/perakende" className="btn">
              <Icon name="frame" size={16} /> Yeni Sipariş
            </Link>
          </>
        }
      />
      <RetailCustomerManager initialQuery={searchParams?.q} />
    </main>
  );
}
