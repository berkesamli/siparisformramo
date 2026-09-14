import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import CekSenetManager from "@/components/CekSenetManager";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function CekSenetPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/ceksenet");
  if (!finansAktif()) redirect("/panel");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1240 }}>
      <PageHeader
        title="Çek / Senet"
        subtitle="Alınan çek cariyi hemen düşürür, kasaya tahsil edildiğinde girer. Ciro edilen çekin kimde olduğu durum sütununda izlenir."
        icon="file-text"
        actions={
          <Link href="/panel/finans" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Finans
          </Link>
        }
      />
      <CekSenetManager />
    </main>
  );
}
