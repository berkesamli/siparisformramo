import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import KasaRaporu from "@/components/KasaRaporu";
import PrintButton from "@/components/PrintButton";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function KasaPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/kasa");
  if (!finansAktif()) redirect("/panel");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1240 }}>
      <PageHeader
        title="Kasa Raporu"
        subtitle="Tüm giriş/çıkış hareketleri — nakit, banka, döviz ve çek/senet kırılımıyla. Çek tahsilleri bankaya tahsil tarihinde işlenir."
        icon="credit-card"
        actions={
          <>
            <Link href="/panel/finans" className="btn secondary no-print">
              <Icon name="chevron-left" size={16} /> Finans
            </Link>
            <PrintButton />
          </>
        }
      />
      <KasaRaporu />
    </main>
  );
}
