import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import Reports from "@/components/Reports";
import PageHeader from "@/components/PageHeader";

export const dynamic = "force-dynamic";

export default async function RaporlarPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/raporlar");
  if (user.role !== "staff") redirect("/portal");
  // Ciro ve tahsilat raporları yalnızca firma sahiplerine açıktır.
  if (!isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <PageHeader
        title="Raporlar"
        subtitle="Ciro, tahsilat, müşteri ve ürün kırılımları — toptan ve perakende birlikte."
        icon="bar-chart"
      />
      {/* Tahsilat/açık bakiye kutuları: tahsilatlar bu sistemde işlenmiyorsa (FINANS_AKTIF kapalı) gizlenir */}
      <Reports finans={finansAktif()} />
    </main>
  );
}
