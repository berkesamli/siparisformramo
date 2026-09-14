import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import FinansDashboard from "@/components/FinansDashboard";
import PageHeader from "@/components/PageHeader";
import Icon, { type IconName } from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function FinansPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans");
  if (!finansAktif()) redirect("/panel");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  const moduller: { href: string; baslik: string; icon: IconName }[] = [
    { href: "/panel/finans/kasa", baslik: "Kasa Raporu", icon: "credit-card" },
    { href: "/panel/finans/giderler", baslik: "Giderler", icon: "arrow-down" },
    { href: "/panel/finans/ceksenet", baslik: "Çek / Senet", icon: "file-text" },
    { href: "/panel/finans/personel", baslik: "Personel", icon: "briefcase" },
    { href: "/panel/raporlar", baslik: "Raporlar", icon: "bar-chart" },
  ];

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <PageHeader
        title="Finans"
        subtitle="Kasa özeti, aylık tahsilat/gider ve vadesi yaklaşan çekler — şube filtresiyle."
        icon="wallet"
      >
        <div className="row no-print" style={{ marginTop: 12 }}>
          {moduller.map((m) => (
            <Link key={m.href} href={m.href} className="btn small secondary">
              <Icon name={m.icon} size={15} /> {m.baslik}
            </Link>
          ))}
        </div>
      </PageHeader>
      <FinansDashboard />
    </main>
  );
}
