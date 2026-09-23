import type { MetadataRoute } from "next";

// Telefonda "Ana ekrana ekle" ile uygulama gibi açılır (tam ekran, açık tema rengi).
export default function manifest(): MetadataRoute.Manifest {
  return {
    name: "Olga Çerçeve — Yönetim Paneli",
    short_name: "Olga Panel",
    description: "Sipariş, stok, fiyat listesi ve kataloglar",
    start_url: "/",
    display: "standalone",
    background_color: "#f4f1ea",
    theme_color: "#f4f1ea",
    lang: "tr",
    icons: [
      { src: "/icon-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
      { src: "/icon-1024.png", sizes: "1024x1024", type: "image/png", purpose: "any" },
      { src: "/logo.png", sizes: "any", type: "image/png", purpose: "any" },
    ],
  };
}
