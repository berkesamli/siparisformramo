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
import { getSessionUser } from "@/lib/auth";
import { isOwner, isFinance, finansAktif, isMaliyet, isKurYetkili } from "@/data/users";
import AppShell from "@/components/shell/AppShell";

export const metadata: Metadata = {
  title: "Olga Çerçeve — Yönetim Paneli",
  description:
    "Olga Çerçeve toptan fiyat listesi, ürün kataloğu, stok durumu ve sipariş sistemi",
  applicationName: "Olga Çerçeve",
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  viewportFit: "cover",
  themeColor: [
    { media: "(prefers-color-scheme: dark)", color: "#0b1120" },
    { media: "(prefers-color-scheme: light)", color: "#f4f1ea" },
  ],
};

// Tema ve kenar çubuğu tercihi ilk boyamadan ÖNCE uygulanır (flash olmasın).
const INIT_SCRIPT = `(function(){try{var t=localStorage.getItem("olga-theme");if(t==="light"||t==="dark"){document.documentElement.setAttribute("data-theme",t)}var s=localStorage.getItem("olga-sb");if(s==="rail"){document.documentElement.setAttribute("data-sb","rail")}}catch(e){}})();`;

export default async function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  const user = await getSessionUser();

  return (
    <html lang="tr" suppressHydrationWarning>
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
            }}
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
