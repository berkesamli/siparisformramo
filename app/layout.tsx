import type { Metadata, Viewport } from "next";
import "./styles/tokens.css";
import "./styles/base.css";
import "./styles/shell.css";
import "./styles/wizard.css";
import "./styles/labels.css";
import "./styles/pickers.css";
import "./styles/orders.css";
import "./styles/reports.css";
import "./styles/modals.css";
import "./styles/dashboard.css";
import "./styles/customers.css";
import "./styles/inbox.css";
import { getSessionUser } from "@/lib/auth";
import { isOwner, isFinance, finansAktif, isMaliyet, isKurYetkili, isMesajci } from "@/data/users";
import AppShell from "@/components/shell/AppShell";
import { aktifDuyurular } from "@/lib/duyurular";
import { istanbulDateKey } from "@/lib/orders";

export const metadata: Metadata = {
  title: "Olga Çerçeve — Yönetim Paneli",
  description:
    "Olga Çerçeve toptan fiyat listesi, ürün kataloğu, stok durumu ve sipariş sistemi",
  applicationName: "Olga Çerçeve",
  icons: { icon: "/logo.png", apple: "/logo.png" },
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  viewportFit: "cover",
  themeColor: "#f4f1ea",
};

// Varsayılan tema AÇIK; kullanıcı koyuyu seçtiyse kayıtlı tercih ilk boyamadan
// ÖNCE uygulanır (flash olmasın). Kenar çubuğu daraltma tercihi de burada okunur.
const INIT_SCRIPT = `(function(){try{var t=localStorage.getItem("olga-theme");if(t==="light"||t==="dark"){document.documentElement.setAttribute("data-theme",t)}var s=localStorage.getItem("olga-sb");if(s==="rail"){document.documentElement.setAttribute("data-sb","rail")}}catch(e){}})();`;

export default async function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  const user = await getSessionUser();
  // Bugün (İstanbul) aktif olan duyurular — rolüne göre süzülür
  const duyurular = user ? aktifDuyurular(user.role, istanbulDateKey()) : [];

  return (
    <html lang="tr" data-theme="light" suppressHydrationWarning>
      <head>
        <script dangerouslySetInnerHTML={{ __html: INIT_SCRIPT }} />
      </head>
      <body>
        {/* Giriş yapılmadan kabuk (menü/üst çubuk) görünmez — ziyaretçi yalnızca
            giriş ekranını görür (sayfalar ayrıca kendi kontrolünü yapar). */}
        {user ? (
          <AppShell
            user={{
              name: user.name,
              username: user.username,
              role: user.role,
              owner: isOwner(user.username),
              finance: finansAktif() && isFinance(user.username),
              maliyet: isMaliyet(user.username),
              kur: isKurYetkili(user.username),
              raporlar: isFinance(user.username),
              mesaj: user.role === "staff" && isMesajci(user.username),
            }}
            duyurular={duyurular}
          >
            {children}
          </AppShell>
        ) : (
          <div className="app-noshell">{children}</div>
        )}
      </body>
    </html>
  );
}
