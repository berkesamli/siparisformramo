import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import PageHeader from "@/components/PageHeader";
import BildirimAyarlari from "@/components/BildirimAyarlari";
import MikroAyarlari from "@/components/MikroAyarlari";
import MesajAyarlari from "@/components/MesajAyarlari";
import IkasAyarlari from "@/components/IkasAyarlari";

export const dynamic = "force-dynamic";

// Bildirim ayarları — yalnızca firma sahipleri.
export default async function AyarlarPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/ayarlar");
  if (user.role !== "staff" || !isOwner(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1000 }}>
      <PageHeader
        icon="bell"
        kicker="Yönetim"
        title="Bildirim Ayarları"
        subtitle="WhatsApp fiş bildirimleri, Mikro bağlantısı, Mesajlar (gelen kutusu) ve ikas (üretim takvimi): kurulum durumu ve test."
      />
      <BildirimAyarlari />
      <MikroAyarlari />
      <MesajAyarlari />
      <IkasAyarlari />
    </main>
  );
}
