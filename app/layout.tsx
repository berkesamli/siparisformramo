import type { Metadata } from "next";
import "./globals.css";
import { getSessionUser } from "@/lib/auth";
import { isOwner, isFinance, finansAktif, isMaliyet, isKurYetkili } from "@/data/users";
import NavBar from "@/components/NavBar";
import NewOrderAlert from "@/components/NewOrderAlert";

export const metadata: Metadata = {
  title: "Olga Çerçeve — Sipariş ve Katalog Platformu",
  description:
    "Olga Çerçeve toptan fiyat listesi, ürün kataloğu, stok durumu ve sipariş sistemi",
};

export default async function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  const user = await getSessionUser();

  return (
    <html lang="tr">
      <body>
        {/* Giriş yapılmadan hiçbir menü görünmez — ziyaretçi yalnızca
            giriş ekranını görür (sayfalar ayrıca kendi kontrolünü yapar). */}
        {user && (
          <NavBar
            user={{
              name: user.name,
              role: user.role,
              owner: isOwner(user.username),
              finance: finansAktif() && isFinance(user.username),
              maliyet: isMaliyet(user.username),
              kur: isKurYetkili(user.username),
            }}
          />
        )}
        {/* Yeni sipariş zili — yalnızca çalışanlarda, tüm panel sayfalarında */}
        {user?.role === "staff" && <NewOrderAlert />}
        {children}
      </body>
    </html>
  );
}
