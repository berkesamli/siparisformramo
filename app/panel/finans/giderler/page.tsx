import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import GiderManager from "@/components/GiderManager";

export const dynamic = "force-dynamic";

export default async function GiderlerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans/giderler");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 260 }}>
          <h1>Giderler</h1>
          <p className="subtitle">Kasa çıkışları — kategori, şube ve yönteme göre.</p>
        </div>
        <Link href="/panel/finans" className="btn small secondary">← Finans</Link>
      </div>
      <GiderManager />
    </main>
  );
}
