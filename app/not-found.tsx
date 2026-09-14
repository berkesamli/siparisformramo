import Link from "next/link";
import Icon from "@/components/shell/Icon";

export default function NotFound() {
  return (
    <main className="container" style={{ maxWidth: 640 }}>
      <div className="card" style={{ textAlign: "center", padding: "44px 24px" }}>
        <div className="empty-icon" style={{ width: 64, height: 64, marginBottom: 14 }}>
          <Icon name="search" size={28} />
        </div>
        <h1 style={{ fontSize: 24 }}>Sayfa bulunamadı</h1>
        <p className="subtitle" style={{ margin: "6px auto 20px" }}>
          Aradığınız sayfa taşınmış ya da hiç var olmamış olabilir.
        </p>
        <div className="row" style={{ justifyContent: "center" }}>
          <Link href="/" className="btn">
            <Icon name="home" size={16} /> Gösterge Paneli
          </Link>
          <Link href="/kataloglar" className="btn secondary">
            <Icon name="book" size={16} /> Kataloglar
          </Link>
        </div>
      </div>
    </main>
  );
}
