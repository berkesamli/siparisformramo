import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";

export const dynamic = "force-dynamic";

// Finans giriş sayfası — Faz 3'te dashboard (kasa özeti, kâr, vade uyarıları)
// buraya gelecek; şimdilik modül kapıları.
export default async function FinansPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/finans");
  if (user.role !== "staff" || !isFinance(user.username)) redirect("/panel");

  const moduller = [
    { href: "/panel/finans/giderler", baslik: "💸 Giderler", aciklama: "Kira, elektrik, maaş, malzeme… kategorili gider kaydı ve aylık döküm." },
    { href: "/panel/finans/ceksenet", baslik: "🧾 Çek / Senet", aciklama: "Portföy, vade uyarıları, tahsil/ciro/karşılıksız takibi." },
    { href: "/panel/finans/personel", baslik: "👥 Personel", aciklama: "Avans, maaş ve prim ödemeleri — kişi bazlı aylık takip." },
    { href: "/panel/raporlar", baslik: "📊 Raporlar", aciklama: "Ciro, tahsilat ve kırılımlar." },
  ];

  return (
    <main className="container">
      <h1>Finans</h1>
      <p className="subtitle">
        Kasa, gider, çek/senet ve personel — tek panelden. Kasa raporu ve genel
        bakış ekranı bir sonraki aşamada bu sayfaya eklenecek.
      </p>
      <div className="cari-cards">
        {moduller.map((m) => (
          <Link key={m.href} href={m.href} className="cari-card" style={{ textDecoration: "none" }}>
            <strong style={{ fontSize: 17 }}>{m.baslik}</strong>
            <span style={{ fontSize: 13 }}>{m.aciklama}</span>
          </Link>
        ))}
      </div>
    </main>
  );
}
