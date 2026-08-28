import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isKurYetkili } from "@/data/users";
import GunlukKur from "@/components/GunlukKur";

export const dynamic = "force-dynamic";

export default async function GunlukKurPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/kur");
  if (user.role !== "staff") redirect("/portal");
  // Günlük kuru yalnızca firma sahipleri belirler
  if (!isKurYetkili(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 820 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ flex: 1 }}>
          <h1>Günlük Kur</h1>
          <p className="subtitle">
            Günün dolar ve euro kurunu belirleyin — bütün sipariş formlarına
            otomatik gelir, çalışanlar değiştiremez.
          </p>
        </div>
        <Link href="/panel" className="btn small secondary">
          ← Sipariş Paneli
        </Link>
      </div>
      <GunlukKur />
    </main>
  );
}
