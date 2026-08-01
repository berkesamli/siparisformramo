import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import StockUpload from "@/components/StockUpload";
import StockSearch from "@/components/StockSearch";

export default async function StockAdminPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/stok");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ flex: 1 }}>
          <h1>Günlük Stok Güncelleme</h1>
          <p className="subtitle">
            Her gün güncel stok Excel&apos;ini yükleyin — müşteri portalındaki
            stok sorgusu anında güncellenir.
          </p>
        </div>
        <Link href="/panel" className="btn small secondary">
          ← Sipariş Paneli
        </Link>
      </div>

      <StockUpload />

      <h2>Yayındaki Stok — Kontrol Edin</h2>
      <div className="card">
        <StockSearch />
      </div>
    </main>
  );
}
