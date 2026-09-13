import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import LabelManager from "@/components/LabelManager";
import PageHeader from "@/components/PageHeader";

export const dynamic = "force-dynamic";

export default async function EtiketPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/etiket");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1280 }}>
      <PageHeader
        title="Müşteriler & Kargo Etiketi"
        subtitle="Müşteri bilgilerini kaydedin, şehre göre listeleyin ve 150×100 mm kargo etiketi yazdırın. Kayıtlı müşteriler sipariş formlarında da seçilebilir."
        icon="users"
      />
      <LabelManager />
    </main>
  );
}
