import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import CustomerDirectory from "@/components/CustomerDirectory";
import PageHeader from "@/components/PageHeader";

export const dynamic = "force-dynamic";

// Müşteri defteri: arama, filtre, hızlı kart erişimi. Etiket yazdırma /etiket'te.
export default async function MusterilerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/musteriler");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1180 }}>
      <PageHeader
        title="Müşteriler"
        subtitle="Kayıtlı bayi ve müşteriler. Bir satıra dokununca müşteri kartı açılır: sipariş geçmişi, Mikro bakiyesi, hızlı işlemler."
        icon="users"
      />
      <CustomerDirectory />
    </main>
  );
}
