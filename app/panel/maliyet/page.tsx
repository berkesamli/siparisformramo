import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { isMaliyet } from "@/data/users";
import MaliyetManager from "@/components/MaliyetManager";

export const dynamic = "force-dynamic";

// Alış fiyatları ve kâr analizi — yalnızca firma sahipleri.
export default async function MaliyetPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/maliyet");
  if (user.role !== "staff" || !isMaliyet(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <h1>Maliyet &amp; Kârlılık</h1>
      <p className="subtitle">
        Ürün alış fiyatları, yüzdesel genel gider ve kod bazlı satış/kâr
        analizi. Parti (konteyner) bazlı: her gelen konteynerin fiyatları ve yüzdesi ayrı girilir.
      </p>
      <MaliyetManager />
    </main>
  );
}
