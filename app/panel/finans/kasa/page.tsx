import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import KasaRaporu from "@/components/KasaRaporu";
import PrintButton from "@/components/PrintButton";

export const dynamic = "force-dynamic";

export default async function KasaPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/kasa");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1240 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 260 }}>
          <h1>Kasa Raporu</h1>
          <p className="subtitle">
            Tüm giriş/çıkış hareketleri — nakit, banka, döviz ve çek/senet
            kırılımıyla. Çek tahsilleri bankaya tahsil tarihinde işlenir.
          </p>
        </div>
        <Link href="/panel/finans" className="btn small secondary no-print">← Finans</Link>
        <PrintButton />
      </div>
      <KasaRaporu />
    </main>
  );
}
