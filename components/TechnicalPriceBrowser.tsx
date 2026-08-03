"use client";

// Teknik malzeme fiyat kataloğu — markaya göre gruplu, kendi arama kutusuyla.
// Arama boşluk/tire duyarsızdır: "cass7kart" → "Cassese 7'lik Agraf Kartuşlu".

import { useMemo, useState } from "react";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { eslesir } from "@/lib/search-norm";

const fmt = (n: number) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const techPrice = (t: (typeof TECHNICAL_PRODUCTS)[number]): string =>
  t.priceTL != null ? `₺${fmt(t.priceTL)}` : `€${fmt(t.priceEUR || 0)}`;

// Bilinen marka sırası; listede olmayan yeni bir kategori eklenirse sessizce
// kaybolmasın diye sonuna eklenir.
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

export default function TechnicalPriceBrowser() {
  const [query, setQuery] = useState("");
  const araniyor = query.trim().length > 0;

  const sonuclar = useMemo(
    () =>
      !araniyor
        ? TECHNICAL_PRODUCTS
        : TECHNICAL_PRODUCTS.filter((t) =>
            eslesir(query, t.name, t.code, t.category)
          ),
    [query, araniyor]
  );

  return (
    <>
      <div className="card no-print" style={{ marginBottom: 20 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Ara — malzeme adı, kod veya marka (askı teli, cass7kart, pozzi)…"
          aria-label="Teknik malzemelerde ara"
        />
        <p style={{ margin: "10px 0 0", color: "var(--muted)", fontSize: 13 }}>
          {araniyor ? (
            sonuclar.length ? (
              <>
                <strong>{sonuclar.length}</strong> malzeme bulundu.
              </>
            ) : (
              <>Sonuç bulunamadı — farklı bir isim veya marka deneyin.</>
            )
          ) : (
            <>{TECHNICAL_PRODUCTS.length} teknik malzeme listeleniyor.</>
          )}
        </p>
      </div>

      {araniyor ? (
        sonuclar.length > 0 && (
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
                  {sonuclar.map((t) => (
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
        )
      ) : (
        CATEGORY_ORDER.map((cat) => {
          const items = TECHNICAL_PRODUCTS.filter((t) => t.category === cat);
          if (!items.length) return null;
          return (
            <div className="card" key={cat} style={{ marginBottom: 20 }}>
              <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>{cat}</h2>
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
  );
}
