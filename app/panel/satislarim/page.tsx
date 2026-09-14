import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import PageHeader from "@/components/PageHeader";
import MySales from "@/components/sales/MySales";

export const dynamic = "force-dynamic";

// Çalışanın kendi satışları: yalnızca oturumdaki kullanıcının adıyla girilen
// siparişler. Veri istemcide /api/cirom'dan gelir (ad oturumdan alınır).
export default async function SatislarimPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/satislarim");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <PageHeader
        icon="trending-up"
        kicker="Kişisel"
        title="Satışlarım"
        subtitle="Yalnızca senin adına girilen siparişler — toptan ve perakende birlikte, iptaller hariç."
      />
      <MySales employeeName={user.name} />
    </main>
  );
}
