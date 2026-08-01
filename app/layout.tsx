import type { Metadata } from "next";
import "./globals.css";
import { getSessionUser } from "@/lib/auth";
import NavBar from "@/components/NavBar";

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
        <NavBar user={user ? { name: user.name, role: user.role } : null} />
        {children}
      </body>
    </html>
  );
}
