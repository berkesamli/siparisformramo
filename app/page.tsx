/* eslint-disable @next/next/no-img-element */
import fs from "fs";
import path from "path";
import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";

export const dynamic = "force-dynamic";

// Ana sayfa kart görselleri: public/anasayfa/ klasörüne
// kataloglar.jpg, stok.jpg, siparis.jpg (jpg/png/webp) eklenince otomatik kullanılır.
function cardImage(base: string): string | null {
  for (const ext of ["jpg", "jpeg", "png", "webp"]) {
    const p = path.join(process.cwd(), "public", "anasayfa", `${base}.${ext}`);
    if (fs.existsSync(p)) return `/anasayfa/${base}.${ext}`;
  }
  return null;
}

export default async function HomePage() {
  const user = await getSessionUser();
  // Site tamamen kapalıdır: giriş yapılmadan ana sayfa da görüntülenemez.
  // Giriş sonrası kullanıcı yine ana sayfaya döner.
  if (!user) redirect("/giris?next=/");

  const cards = [
    {
      title: "Kataloglar",
      sub: "PDF · Dergi Görünümü",
      href: "/kataloglar",
      img: cardImage("kataloglar"),
    },
    {
      title: "Stok Durumu",
      sub: "Ankara · İstanbul",
      href: user ? "/portal" : "/giris?next=/portal",
      img: cardImage("stok"),
    },
    {
      title: "Sipariş Paneli",
      sub: "Çalışanlara Özel",
      href: user?.role === "staff" ? "/panel" : "/giris?next=/panel",
      img: cardImage("siparis"),
    },
    {
      title: "Online Çerçeve",
      sub: "Perakende · Çerçeveletme",
      href:
        user?.role === "staff"
          ? "/panel/perakende"
          : "/giris?next=/panel/perakende",
      img: cardImage("perakende"),
    },
  ];

  return (
    <main className="container">
      <div style={{ textAlign: "center", padding: "26px 0 26px" }}>
        <p
          className="subtitle"
          style={{ fontSize: 15.5, maxWidth: 620, margin: "0 auto" }}
        >
          Profesyonel çerçeveleme ve dekorasyon çözümleri — kataloglar, güncel
          stok ve online sipariş platformu.
        </p>
      </div>

      <div className="hero-grid">
        {cards.map((c) => (
          <Link key={c.title} href={c.href} className="hero-card">
            <span className="hero-media">
              {c.img ? (
                <img className="hero-bg" src={c.img} alt={c.title} />
              ) : (
                <img className="hero-watermark" src="/logo.png" alt="" />
              )}
              <span className="hero-shade" />
            </span>
            <span className="hero-txt">
              <h3>{c.title}</h3>
              <span>{c.sub}</span>
            </span>
            <span className="hero-btn" aria-hidden>
              <svg
                width="20"
                height="20"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.4"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <line x1="7" y1="17" x2="17" y2="7" />
                <polyline points="8 7 17 7 17 16" />
              </svg>
            </span>
          </Link>
        ))}
      </div>

      <div
        className="card"
        style={{ marginTop: 34, textAlign: "center", maxWidth: 1080, marginLeft: "auto", marginRight: "auto" }}
      >
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
