import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import FinansDashboard from "@/components/FinansDashboard";

export const dynamic = "force-dynamic";

export default async function FinansPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  const moduller = [
    { href: "/panel/finans/kasa", baslik: "🧮 Kasa Raporu" },
    { href: "/panel/finans/giderler", baslik: "💸 Giderler" },
    { href: "/panel/finans/ceksenet", baslik: "🧾 Çek / Senet" },
    { href: "/panel/finans/personel", baslik: "👥 Personel" },
    { href: "/panel/raporlar", baslik: "📊 Raporlar" },
  ];

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <h1 style={{ flex: 1, minWidth: 200 }}>Finans</h1>
        {moduller.map((m) => (
          <Link key={m.href} href={m.href} className="btn small secondary">
            {m.baslik}
          </Link>
        ))}
      </div>
      <p className="subtitle">
        Kasa özeti, aylık tahsilat/gider ve vadesi yaklaşan çekler — şube
        filtresiyle.
      </p>
      <FinansDashboard />
    </main>
  );
}
