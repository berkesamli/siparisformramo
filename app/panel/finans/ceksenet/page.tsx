import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import CekSenetManager from "@/components/CekSenetManager";

export const dynamic = "force-dynamic";

export default async function CekSenetPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/ceksenet");
  if (!finansAktif()) redirect("/panel");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1240 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 260 }}>
          <h1>Çek / Senet</h1>
          <p className="subtitle">
            Alınan çek cariyi hemen düşürür, kasaya tahsil edildiğinde girer.
            Ciro edilen çekin kimde olduğu durum sütununda izlenir.
          </p>
        </div>
        <Link href="/panel/finans" className="btn small secondary">← Finans</Link>
      </div>
      <CekSenetManager />
    </main>
  );
}
