import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { FRAME_PROFILES, SERIES_ORDER } from "@/data/catalog";
import PrintButton from "@/components/PrintButton";

const fmt = (n: number) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

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
            Profiller koli bazında satılır.
          </p>
        </div>
        <PrintButton />
      </div>

      {SERIES_ORDER.map((series) => {
        const items = FRAME_PROFILES.filter((f) => f.series === series);
        if (!items.length) return null;
        return (
          <div className="card" key={series} style={{ marginBottom: 20 }}>
            <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>
              {series} Serisi
            </h2>
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr>
                    <th>Ürün Kodu</th>
                    <th>Koli Adet</th>
                    <th>Koli Metraj</th>
                    <th>Fiyat (USD/mt)</th>
                  </tr>
                </thead>
                <tbody>
                  {items.map((f) => (
                    <tr key={f.code}>
                      <td style={{ fontWeight: 600 }}>{f.code}</td>
                      <td>{f.koliAdet}</td>
                      <td>{fmt(f.koliMetraj)} MT</td>
                      <td>${fmt(f.priceUSD)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}

      <p style={{ color: "var(--muted)", fontSize: 13 }}>
        Fiyatlar güncellenebilir; en güncel fiyat için 0850 305 75 45 numaralı
        sipariş hattımızla iletişime geçiniz. Satış yapılırken fiyatlara KDV
        ilave edilir.
      </p>
    </main>
  );
}
