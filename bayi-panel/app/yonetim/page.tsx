import { redirect } from "next/navigation";
import { getAdminSession } from "@/lib/auth";
import DealersAdmin from "@/components/DealersAdmin";

export const dynamic = "force-dynamic";

export default async function YonetimPage() {
  const admin = await getAdminSession();
  if (!admin) redirect("/giris?next=/yonetim");
  return (
    <main className="container" style={{ maxWidth: 1100 }}>
      <h1 style={{ marginBottom: 4 }}>Bayi Yönetimi</h1>
      <p className="subtitle">Bayi hesapları, abonelik durumları ve kullanım istatistikleri.</p>
      <DealersAdmin />
    </main>
  );
}
