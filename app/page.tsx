import fs from "fs";
import path from "path";
import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { isFinance, isKurYetkili, finansAktif } from "@/data/users";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";
import Dashboard from "@/components/dashboard/Dashboard";
import CustomerDashboard from "@/components/dashboard/CustomerDashboard";
import AiChat from "@/components/AiChat";
import type { QuickTile } from "@/components/dashboard/QuickActions";

export const dynamic = "force-dynamic";

// Hızlı işlem kartı görselleri: public/anasayfa/ klasöründeki
// kataloglar/stok/siparis/perakende (jpg/png/webp) otomatik kullanılır.
function cardImage(base: string): string | null {
  for (const ext of ["jpg", "jpeg", "png", "webp"]) {
    const p = path.join(process.cwd(), "public", "anasayfa", `${base}.${ext}`);
    if (fs.existsSync(p)) return `/anasayfa/${base}.${ext}`;
  }
  return null;
}

function tarihStr(): string {
  return new Date().toLocaleDateString("tr-TR", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
    timeZone: "Europe/Istanbul",
  });
}

export default async function HomePage() {
  const user = await getSessionUser();
  // Site tamamen kapalıdır: giriş yapılmadan ana sayfa da görüntülenemez.
  if (!user) redirect("/giris?next=/");

  const staff = user.role === "staff";
  const finance = staff && isFinance(user.username);
  const ilkAd = user.name.split(/\s+/)[0] || user.name;

  if (!staff) {
    const tiles: QuickTile[] = [
      { title: "Stok Durumu", sub: "Ankara · İstanbul depoları", href: "/portal", icon: "package", img: cardImage("stok") },
      { title: "Toptan Fiyat Listesi", sub: "Profiller ve teknik malzeme", href: "/portal/fiyat-listesi", icon: "tag", img: cardImage("siparis") },
      { title: "Kataloglar", sub: "PDF · dergi görünümü", href: "/kataloglar", icon: "book", img: cardImage("kataloglar") },
    ];
    return (
      <main className="container">
        <PageHeader
          kicker={tarihStr()}
          title={`Hoş geldiniz, ${ilkAd}`}
          subtitle={"Güncel stok, toptan fiyat listesi ve kataloglara buradan ulaşabilirsiniz. Sipariş için: 0850\u00A0305\u00A075\u00A045"}
          icon="home"
        />
        <CustomerDashboard tiles={tiles} />
        <AiChat />
      </main>
    );
  }

  const tiles: QuickTile[] = [
    { title: "Yeni Toptan Sipariş", sub: "Sipariş formu · e\u2011posta + WhatsApp", href: "/panel", icon: "plus", img: cardImage("siparis") },
    { title: "Online Çerçeve", sub: "Perakende çerçeveletme sihirbazı", href: "/panel/perakende", icon: "frame", img: cardImage("perakende") },
    { title: "Stok Sorgula", sub: "Ankara · İstanbul", href: "/portal", icon: "package", img: cardImage("stok") },
    { title: "Kataloglar", sub: "PDF · dergi görünümü", href: "/kataloglar", icon: "book", img: cardImage("kataloglar") },
  ];
  const chips: { label: string; href: string; icon: "users" | "message" | "upload" | "tag" | "book" | "dollar" | "bar-chart" }[] = [
    { label: "Müşteriler & Etiket", href: "/etiket", icon: "users" },
    { label: "SMS Gönder", href: "/panel/sms", icon: "message" },
    { label: "Stok Yükle", href: "/panel/stok", icon: "upload" },
    { label: "Fiyat Listesi", href: "/portal/fiyat-listesi", icon: "tag" },
  ];
  if (isKurYetkili(user.username)) chips.push({ label: "Günlük Kur", href: "/panel/kur", icon: "dollar" });
  if (finance) chips.push({ label: "Raporlar", href: "/panel/raporlar", icon: "bar-chart" });
  if (finance && finansAktif()) chips.push({ label: "Finans", href: "/panel/finans", icon: "bar-chart" });

  return (
    <main className="container">
      <PageHeader
        kicker={tarihStr()}
        title={`Günün özeti — hoş geldin, ${ilkAd}`}
        subtitle="Bugünkü siparişler, açık işler ve dikkat gerektiren konular tek bakışta."
        icon="home"
        actions={
          <>
            <Link href="/panel/perakende" className="btn secondary">
              <Icon name="frame" size={17} /> Online Çerçeve
            </Link>
            <Link href="/panel" className="btn">
              <Icon name="plus" size={17} /> Yeni Sipariş
            </Link>
          </>
        }
      />
      <Dashboard finance={finance} tiles={tiles} chips={chips} />
      <AiChat />
    </main>
  );
}
