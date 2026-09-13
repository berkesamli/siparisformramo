import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import PersonelManager from "@/components/PersonelManager";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function PersonelPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/personel");
  if (!finansAktif()) redirect("/panel");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1240 }}>
      <PageHeader
        title="Personel"
        subtitle="Avans, maaş ve prim ödemeleri — ödemeler gider kaydına da düşer."
        icon="briefcase"
        actions={
          <Link href="/panel/finans" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Finans
          </Link>
        }
      />
      <PersonelManager />
    </main>
  );
}
