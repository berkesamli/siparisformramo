import type { Metadata } from "next";
import "./globals.css";
import { getSessionUser } from "@/lib/auth";
import NavBar from "@/components/NavBar";

export const metadata: Metadata = {
  title: "Olga Çerçeve — Bayi Paneli",
  description: "Olga Çerçeve bayileri için online çerçeve fiyatlandırma ve sipariş takip paneli",
};

export const dynamic = "force-dynamic";

export default async function RootLayout({ children }: { children: React.ReactNode }) {
  const user = await getSessionUser();
  return (
    <html lang="tr">
      <body>
        {user && <NavBar user={{ kind: user.kind, name: user.name }} />}
        {children}
      </body>
    </html>
  );
}
