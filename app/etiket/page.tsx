import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import LabelManager from "@/components/LabelManager";
import PageHeader from "@/components/PageHeader";

export const dynamic = "force-dynamic";

// Kargo etiketi: kayıtlı müşteriyi seç, 150×100 mm etiketi PDF olarak yazdır.
// Müşteri ekleme/düzenleme ana yeri /musteriler; burada hızlı düzenleme var.
export default async function EtiketPage({ searchParams }: { searchParams: { id?: string } }) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/etiket");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1280 }}>
      <PageHeader
        title="Kargo Etiketi"
        subtitle="Müşteriyi seçin, gönderici şubeyi belirleyin ve 150 × 100 mm etiketi PDF olarak yazdırın."
        icon="tag"
      />
      <LabelManager preselectId={searchParams.id || ""} />
    </main>
  );
}
