import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import PriceListBrowser from "@/components/PriceListBrowser";
import PrintButton from "@/components/PrintButton";

export default async function PriceListPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/portal/fiyat-listesi");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ flex: 1 }}>
          <h1>Toptan Fiyat Listesi</h1>
          <p className="subtitle">
            Çerçeve profilleri — fiyatlar USD/mt cinsinden ve KDV hariçtir.
            Profiller koli bazında satılır. Teknik malzemeler için üstteki
            butonu kullanın.
          </p>
        </div>
        <PrintButton />
      </div>

      <PriceListBrowser />

      <p style={{ color: "var(--muted)", fontSize: 13 }}>
        Fiyatlar güncellenebilir; en güncel fiyat için 0850 305 75 45 numaralı
        sipariş hattımızla iletişime geçiniz. Satış yapılırken fiyatlara KDV
        ilave edilir.
      </p>
    </main>
  );
}
