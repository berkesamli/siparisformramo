import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isKurYetkili } from "@/data/users";
import GunlukKur from "@/components/GunlukKur";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function GunlukKurPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/kur");
  if (user.role !== "staff") redirect("/portal");
  // Günlük kuru yalnızca firma sahipleri belirler
  if (!isKurYetkili(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 820 }}>
      <PageHeader
        title="Günlük Kur"
        subtitle="Günün dolar ve euro kurunu belirleyin — bütün sipariş formlarına otomatik gelir, çalışanlar değiştiremez."
        icon="dollar"
        actions={
          <Link href="/panel" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Sipariş Paneli
          </Link>
        }
      />
      <GunlukKur />
    </main>
  );
}
