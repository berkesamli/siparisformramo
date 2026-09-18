import { redirect } from "next/navigation";

export const dynamic = "force-dynamic";

// Eski adres: /musteri?id=C123 → /musteriler/kart?id=C123 (arama sonuçları, yer imleri)
export default function EskiMusteriPage({ searchParams }: { searchParams: { id?: string } }) {
  const id = searchParams.id || "";
  redirect(id ? `/musteriler/kart?id=${encodeURIComponent(id)}` : "/musteriler");
}
