import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import StockUpload from "@/components/StockUpload";
import StockMikro from "@/components/StockMikro";
import { isOwner } from "@/data/users";
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
        title="Stok Güncelleme"
        subtitle="Stok artık Mikro'dan çekilir; Excel yüklemesi yedek olarak durur. Müşteri portalındaki stok sorgusu anında güncellenir."
        icon="upload"
        actions={
          <Link href="/panel" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Sipariş Paneli
          </Link>
        }
      />

      <StockMikro owner={isOwner(user.username)} />

      <div className="card" style={{ marginTop: 16 }}>
        <div className="card-head">
          <span className="card-head-icon"><Icon name="upload" size={18} /></span>
          <div>
            <h2>Excel ile Yükleme (yedek)</h2>
            <span className="card-head-sub">Mikro bağlantısı çalışmazsa eski yöntem</span>
          </div>
        </div>
        <StockUpload />
      </div>

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
