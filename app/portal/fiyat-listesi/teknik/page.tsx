import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import TechnicalPriceBrowser from "@/components/TechnicalPriceBrowser";
import PrintButton from "@/components/PrintButton";

export default async function TechnicalPriceListPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/portal/fiyat-listesi/teknik");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 260 }}>
          <h1>Teknik Malzeme Fiyat Listesi</h1>
          <p className="subtitle">
            € işaretli ürünler Euro, ₺ işaretli ürünler Türk Lirası üzerinden
            fiyatlandırılır. Fiyatlar kutu bazındadır ve KDV hariçtir.
          </p>
        </div>
        <Link href="/portal/fiyat-listesi" className="btn small secondary no-print">
          ← Çerçeve Fiyatları
        </Link>
        <PrintButton />
      </div>

      <TechnicalPriceBrowser />

      <p style={{ color: "var(--muted)", fontSize: 13 }}>
        Fiyatlar güncellenebilir; en güncel fiyat için 0850 305 75 45 numaralı
        sipariş hattımızla iletişime geçiniz. Satış yapılırken fiyatlara KDV
        ilave edilir.
      </p>
    </main>
  );
}
