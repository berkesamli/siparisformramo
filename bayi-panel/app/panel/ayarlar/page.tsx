import { redirect } from "next/navigation";
import { getDealerSession } from "@/lib/auth";
import PricingSettings from "@/components/PricingSettings";

export const dynamic = "force-dynamic";

export default async function AyarlarPage() {
  const s = await getDealerSession();
  if (!s) redirect("/giris?next=/panel/ayarlar");
  return (
    <main className="container" style={{ maxWidth: 1000 }}>
      <h1 style={{ marginBottom: 4 }}>Fiyat & Ayarlar</h1>
      <p className="subtitle">
        Müşterinize sunduğunuz fiyatları buradan belirlersiniz. Olga toptan liste fiyatları katalogdan otomatik gelir.
      </p>
      <PricingSettings />
    </main>
  );
}
