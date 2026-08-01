import Link from "next/link";
import { getSessionUser } from "@/lib/auth";

export default async function HomePage() {
  const user = await getSessionUser();

  return (
    <main className="container">
      <div style={{ textAlign: "center", padding: "60px 0 40px" }}>
        <h1 style={{ fontSize: 38 }}>
          Olga <span style={{ color: "var(--brand-light)" }}>Çerçeve</span>
        </h1>
        <p className="subtitle" style={{ fontSize: 17, maxWidth: 640, margin: "10px auto 0" }}>
          Profesyonel çerçeveleme ve dekorasyon çözümleri — toptan fiyat listesi,
          ürün katalogları, stok durumu ve online sipariş platformu.
        </p>
      </div>

      <div className="grid cols-3">
        <div className="card">
          <h2 style={{ marginTop: 0 }}>📖 Kataloglar</h2>
          <p style={{ color: "var(--text-2)", marginBottom: 16 }}>
            Çerçeve profili ve teknik malzeme kataloglarımızı dergi formatında
            inceleyin.
          </p>
          <Link href="/kataloglar" className="btn">
            Katalogları Aç
          </Link>
        </div>

        <div className="card">
          <h2 style={{ marginTop: 0 }}>🏷️ Ürünler &amp; Stok</h2>
          <p style={{ color: "var(--text-2)", marginBottom: 16 }}>
            Bayilerimiz için: tüm serilerde stok durumu ve toptan fiyat listesi.
            Giriş gereklidir.
          </p>
          <Link href={user ? "/portal" : "/giris?next=/portal"} className="btn">
            Portala Git
          </Link>
        </div>

        <div className="card">
          <h2 style={{ marginTop: 0 }}>🧾 Sipariş Paneli</h2>
          <p style={{ color: "var(--text-2)", marginBottom: 16 }}>
            Çalışanlarımız için: sipariş oluşturma, e-posta ve WhatsApp ile
            iletim.
          </p>
          <Link href={user?.role === "staff" ? "/panel" : "/giris?next=/panel"} className="btn">
            Panele Git
          </Link>
        </div>
      </div>

      <div className="card" style={{ marginTop: 24, textAlign: "center" }}>
        <p style={{ color: "var(--text-2)" }}>
          📞 Sipariş Hattı: <strong>0850 305 75 45</strong> · Ankara: 0312 495 75 45 ·
          İstanbul: 0212 675 27 50 ·{" "}
          <a href="https://olgacerceve.com" target="_blank" rel="noreferrer">
            olgacerceve.com
          </a>
        </p>
      </div>
    </main>
  );
}
