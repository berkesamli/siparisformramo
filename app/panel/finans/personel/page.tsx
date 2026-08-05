import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance, finansAktif } from "@/data/users";
import PersonelManager from "@/components/PersonelManager";

export const dynamic = "force-dynamic";

export default async function PersonelPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/personel");
  if (!finansAktif()) redirect("/panel");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1100 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 260 }}>
          <h1>Personel</h1>
          <p className="subtitle">Avans, maaş ve prim ödemeleri — ödemeler gider kaydına da düşer.</p>
        </div>
        <Link href="/panel/finans" className="btn small secondary">← Finans</Link>
      </div>
      <PersonelManager />
    </main>
  );
}
