import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import StockUpload from "@/components/StockUpload";
import StockSearch from "@/components/StockSearch";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export default async function StockAdminPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/stok");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <PageHeader
        title="Günlük Stok Güncelleme"
        subtitle="Her gün güncel stok Excel'ini yükleyin — müşteri portalındaki stok sorgusu anında güncellenir."
        icon="upload"
        actions={
          <Link href="/panel" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Sipariş Paneli
          </Link>
        }
      />

      <StockUpload />

      <div className="card" style={{ marginTop: 16 }}>
        <div className="card-head">
          <span className="card-head-icon">
            <Icon name="search" size={18} />
          </span>
          <div>
            <h2>Yayındaki Stok — Kontrol Edin</h2>
          </div>
        </div>
        <StockSearch />
      </div>
    </main>
  );
}
