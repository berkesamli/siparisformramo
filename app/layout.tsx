import type { Metadata } from "next";
import "./globals.css";
import { getSessionUser } from "@/lib/auth";
import Link from "next/link";
import LogoutButton from "@/components/LogoutButton";

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
        <nav className="topnav">
          <Link href="/" className="logo">
            OLGA ÇERÇEVE
          </Link>
          <Link href="/kataloglar" className="navlink">
            Kataloglar
          </Link>
          {user && (
            <Link href="/portal" className="navlink">
              Ürünler &amp; Stok
            </Link>
          )}
          {user && (
            <Link href="/portal/fiyat-listesi" className="navlink">
              Toptan Fiyat Listesi
            </Link>
          )}
          {user?.role === "staff" && (
            <Link href="/panel" className="navlink">
              Sipariş Paneli
            </Link>
          )}
          {user?.role === "staff" && (
            <Link href="/panel/siparisler" className="navlink">
              Siparişler
            </Link>
          )}
          {user?.role === "staff" && (
            <Link href="/panel/stok" className="navlink">
              Stok Yükle
            </Link>
          )}
          <span className="spacer" />
          {user ? (
            <>
              <span className="user">
                {user.name} ({user.role === "staff" ? "Çalışan" : "Müşteri"})
              </span>
              <LogoutButton />
            </>
          ) : (
            <Link href="/giris" className="btn small">
              Giriş Yap
            </Link>
          )}
        </nav>
        {children}
      </body>
    </html>
  );
}
