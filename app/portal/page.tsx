import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import ProductBrowser from "@/components/ProductBrowser";
import AiChat from "@/components/AiChat";

export default async function PortalPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/portal");

  return (
    <main className="container">
      <h1>Ürünler &amp; Stok Durumu</h1>
      <p className="subtitle">
        Çerçeve profilleri ve teknik malzemeler — toptan liste fiyatları KDV
        hariçtir. Güncel stok için sipariş hattımızı arayabilirsiniz: 0850 305 75 45
      </p>
      <ProductBrowser />
      <AiChat />
    </main>
  );
}
