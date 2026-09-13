import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import GiderManager from "@/components/GiderManager";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function GiderlerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/giderler");
  if (!finansAktif()) redirect("/panel");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <PageHeader
        title="Giderler"
        subtitle="Kasa çıkışları — kategori, şube ve yönteme göre."
        icon="arrow-down"
        actions={
          <Link href="/panel/finans" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Finans
          </Link>
        }
      />
      <GiderManager />
    </main>
  );
}
