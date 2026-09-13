import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import StockSearch from "@/components/StockSearch";
import AiChat from "@/components/AiChat";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export default async function PortalPage({
  searchParams,
}: {
  searchParams?: { q?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/portal");

  return (
    <main className="container">
      <PageHeader
        title="Güncel Stok Sorgulama"
        subtitle="Ürün kodunu yazın — Ankara ve İstanbul depolarındaki güncel stok anında listelenir. Sipariş hattı: 0850 305 75 45"
        icon="package"
        actions={
          <Link href="/portal/fiyat-listesi" className="btn">
            <Icon name="tag" size={16} /> Fiyat Listesi
            <Icon name="chevron-right" size={15} />
          </Link>
        }
      />

      <div className="card">
        <StockSearch initialQuery={searchParams?.q} />
      </div>

      <AiChat />
    </main>
  );
}
