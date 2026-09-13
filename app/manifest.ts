import type { MetadataRoute } from "next";

// Telefonda "Ana ekrana ekle" ile uygulama gibi açılır (tam ekran, koyu tema rengi).
export default function manifest(): MetadataRoute.Manifest {
  return {
    name: "Olga Çerçeve — Yönetim Paneli",
    short_name: "Olga Panel",
    description: "Sipariş, stok, fiyat listesi ve kataloglar",
    start_url: "/",
    display: "standalone",
    background_color: "#0b1120",
    theme_color: "#0b1120",
    lang: "tr",
    icons: [
      { src: "/logo.png", sizes: "any", type: "image/png", purpose: "any" },
    ],
  };
}
