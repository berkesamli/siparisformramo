import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import StockSearch from "@/components/StockSearch";
import AiChat from "@/components/AiChat";

export default async function PortalPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/portal");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 260 }}>
          <h1>Güncel Stok Sorgulama</h1>
          <p className="subtitle">
            Ürün kodunu yazın — Ankara ve İstanbul depolarındaki güncel stok
            anında listelenir. Sipariş hattı: 0850 305 75 45
          </p>
        </div>
        <Link href="/portal/fiyat-listesi" className="btn small secondary">
          Fiyat Listesi →
        </Link>
      </div>

      <div className="card">
        <StockSearch />
      </div>

      <AiChat />
    </main>
  );
}
