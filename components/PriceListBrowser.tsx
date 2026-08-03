"use client";

// Toptan fiyat listesi — çerçeve profilleri ve teknik malzemeler tek sayfada,
// tek arama kutusuyla. Arama Türkçe karakter ve büyük/küçük harf duyarsızdır:
// "aski", "ASKI", "askı" aynı sonucu verir.

import { useMemo, useState } from "react";
import { FRAME_PROFILES, SERIES_ORDER } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";

const fmt = (n: number) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const TR_TO_ASCII: Record<string, string> = {
  ç: "c", Ç: "c", ğ: "g", Ğ: "g", ı: "i", İ: "i",
  ö: "o", Ö: "o", ş: "s", Ş: "s", ü: "u", Ü: "u",
};

function norm(s: string): string {
  return String(s || "")
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_TO_ASCII[c])
    .toLowerCase()
    .trim();
}

// Teknik malzemeler markaya göre gruplanır; sıra fiyat listesindeki düzeni takip
// eder. Listede olmayan yeni bir kategori eklenirse sessizce kaybolmasın diye
// bilinen sıranın ardına eklenir.
const BILINEN_SIRA = [
  "Pozzi",
  "Alfamacchine",
  "Cassese",
  "Danlist",
  "Ro-ma Maestri",
  "Scappi Cartoni",
  "OLGA",
  "NS Serisi",
];

const CATEGORY_ORDER = [
  ...BILINEN_SIRA,
  ...Array.from(new Set(TECHNICAL_PRODUCTS.map((t) => t.category))).filter(
    (c) => !BILINEN_SIRA.includes(c)
  ),
];

function techPrice(t: (typeof TECHNICAL_PRODUCTS)[number]): string {
  return t.priceTL != null ? `₺${fmt(t.priceTL)}` : `€${fmt(t.priceEUR || 0)}`;
}

export default function PriceListBrowser() {
  const [query, setQuery] = useState("");
  const q = norm(query);

  const frames = useMemo(
    () =>
      !q
        ? FRAME_PROFILES
        : FRAME_PROFILES.filter(
            (f) => norm(f.code).includes(q) || norm(f.series).includes(q)
          ),
    [q]
  );

  const technicals = useMemo(
    () =>
      !q
        ? TECHNICAL_PRODUCTS
        : TECHNICAL_PRODUCTS.filter(
            (t) =>
              norm(t.name).includes(q) ||
              norm(t.code).includes(q) ||
              norm(t.category).includes(q)
          ),
    [q]
  );

  const araniyor = q.length > 0;
  const toplam = frames.length + technicals.length;

  return (
    <>
      <div className="card no-print" style={{ marginBottom: 20 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Ara — çerçeve kodu (GB022) veya malzeme adı (askı teli, agraf, karton)…"
          aria-label="Fiyat listesinde ara"
        />
        <p style={{ margin: "10px 0 0", color: "var(--muted)", fontSize: 13 }}>
          {araniyor ? (
            toplam ? (
              <>
                <strong>{frames.length}</strong> çerçeve profili,{" "}
                <strong>{technicals.length}</strong> teknik malzeme bulundu.
              </>
            ) : (
              <>Sonuç bulunamadı — farklı bir kod veya isim deneyin.</>
            )
          ) : (
            <>
              {FRAME_PROFILES.length} çerçeve profili ve{" "}
              {TECHNICAL_PRODUCTS.length} teknik malzeme listeleniyor.
            </>
          )}
        </p>
      </div>

      {/* ---------- Çerçeve profilleri ---------- */}
      {araniyor ? (
        frames.length > 0 && (
          <div className="card" style={{ marginBottom: 20 }}>
            <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>
              Çerçeve Profilleri ({frames.length})
            </h2>
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr>
                    <th>Ürün Kodu</th>
                    <th>Seri</th>
                    <th>Koli Adet</th>
                    <th>Koli Metraj</th>
                    <th>Fiyat (USD/mt)</th>
                  </tr>
                </thead>
                <tbody>
                  {frames.map((f) => (
                    <tr key={f.code}>
                      <td style={{ fontWeight: 600 }}>{f.code}</td>
                      <td>{f.series}</td>
                      <td>{f.koliAdet}</td>
                      <td>{fmt(f.koliMetraj)} MT</td>
                      <td>${fmt(f.priceUSD)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )
      ) : (
        SERIES_ORDER.map((series) => {
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
        })
      )}

      {/* ---------- Teknik malzemeler ---------- */}
      {technicals.length > 0 && (
        <>
          <h2 style={{ color: "var(--brand-light)", marginTop: 28 }}>
            Teknik Malzemeler{araniyor ? ` (${technicals.length})` : ""}
          </h2>
          <p className="subtitle" style={{ marginTop: -8 }}>
            € işaretli ürünler Euro, ₺ işaretli ürünler Türk Lirası üzerinden
            fiyatlandırılır. Fiyatlar kutu bazındadır ve KDV hariçtir.
          </p>

          {araniyor ? (
            <div className="card" style={{ marginBottom: 20 }}>
              <div style={{ overflowX: "auto" }}>
                <table>
                  <thead>
                    <tr>
                      <th>Ürün</th>
                      <th>Marka / Kategori</th>
                      <th>Adet / Kutu</th>
                      <th>Fiyat</th>
                    </tr>
                  </thead>
                  <tbody>
                    {technicals.map((t) => (
                      <tr key={t.code}>
                        <td style={{ fontWeight: 600 }}>{t.name}</td>
                        <td>{t.category}</td>
                        <td>{t.adetPerKutu.toLocaleString("tr-TR")}</td>
                        <td>{techPrice(t)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            CATEGORY_ORDER.map((cat) => {
              const items = TECHNICAL_PRODUCTS.filter((t) => t.category === cat);
              if (!items.length) return null;
              return (
                <div className="card" key={cat} style={{ marginBottom: 20 }}>
                  <h3 style={{ marginTop: 0 }}>{cat}</h3>
                  <div style={{ overflowX: "auto" }}>
                    <table>
                      <thead>
                        <tr>
                          <th>Ürün</th>
                          <th>Adet / Kutu</th>
                          <th>Fiyat</th>
                        </tr>
                      </thead>
                      <tbody>
                        {items.map((t) => (
                          <tr key={t.code}>
                            <td style={{ fontWeight: 600 }}>{t.name}</td>
                            <td>{t.adetPerKutu.toLocaleString("tr-TR")}</td>
                            <td>{techPrice(t)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })
          )}
        </>
      )}
    </>
  );
}
