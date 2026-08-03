"use client";

// Toptan fiyat listesi — çerçeve profilleri seri seri listelenir, teknik
// malzemeler ayrı sayfada durur (üstteki butonla yeni sekmede açılır).
// Arama kutusu ikisini birden tarar: bir malzeme arandığında sonucu bu sayfada
// da gösterilir, kullanıcı diğer sayfaya geçmek zorunda kalmaz.

import { useMemo, useState } from "react";
import { FRAME_PROFILES, SERIES_ORDER } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { eslesir } from "@/lib/search-norm";

const fmt = (n: number) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const techPrice = (t: (typeof TECHNICAL_PRODUCTS)[number]): string =>
  t.priceTL != null ? `₺${fmt(t.priceTL)}` : `€${fmt(t.priceEUR || 0)}`;

export default function PriceListBrowser() {
  const [query, setQuery] = useState("");
  const araniyor = query.trim().length > 0;

  const frames = useMemo(
    () =>
      !araniyor
        ? FRAME_PROFILES
        : FRAME_PROFILES.filter((f) => eslesir(query, f.code, f.series)),
    [query, araniyor]
  );

  const technicals = useMemo(
    () =>
      !araniyor
        ? []
        : TECHNICAL_PRODUCTS.filter((t) =>
            eslesir(query, t.name, t.code, t.category)
          ),
    [query, araniyor]
  );

  const toplam = frames.length + technicals.length;

  return (
    <>
      <div className="card no-print" style={{ marginBottom: 20 }}>
        <div
          style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}
        >
          <input
            style={{ flex: 1, minWidth: 240 }}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Ara — çerçeve kodu (2315S, GB022) veya malzeme adı (askı teli, agraf)…"
            aria-label="Fiyat listesinde ara"
          />
          <a
            href="/portal/fiyat-listesi/teknik"
            target="_blank"
            rel="noreferrer"
            className="btn small secondary"
          >
            🔧 Teknik Malzemeler ↗
          </a>
        </div>

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
              {FRAME_PROFILES.length} çerçeve profili listeleniyor. Teknik
              malzemeler için sağdaki butonu kullanın; arama kutusu ikisini
              birden tarar.
            </>
          )}
        </p>
      </div>

      {/* ---------- Teknik malzeme sonuçları (arama yapılınca üstte) ---------- */}
      {araniyor && technicals.length > 0 && (
        <div className="card" style={{ marginBottom: 20 }}>
          <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>
            Teknik Malzemeler ({technicals.length})
          </h2>
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
      )}

      {/* ---------- Çerçeve profilleri ---------- */}
      {araniyor
        ? frames.length > 0 && (
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
        : SERIES_ORDER.map((series) => {
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
    </>
  );
}
