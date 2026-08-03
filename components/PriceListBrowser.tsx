"use client";

// Toptan fiyat listesi — arama kutusunun altındaki iki kutu ile çerçeve
// profilleri ve teknik malzemeler arasında geçiş yapılır. Aynı sayfada kalınır;
// teknik malzemelere ulaşmak için aşağı kaydırmak gerekmez.
//
// Arama yapıldığında sekme farkı gözetilmez: her iki gruptan eşleşen ürünler
// birlikte listelenir, kullanıcı hangi sekmede olduğunu düşünmek zorunda kalmaz.

import { useMemo, useState } from "react";
import { FRAME_PROFILES, SERIES_ORDER } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { eslesir } from "@/lib/search-norm";

const fmt = (n: number) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const techPrice = (t: (typeof TECHNICAL_PRODUCTS)[number]): string =>
  t.priceTL != null ? `₺${fmt(t.priceTL)}` : `€${fmt(t.priceEUR || 0)}`;

// Bilinen marka sırası; listede olmayan yeni bir kategori eklenirse sessizce
// kaybolmasın diye sıranın sonuna eklenir.
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

function FrameTable({
  items,
  seriGoster,
}: {
  items: typeof FRAME_PROFILES;
  seriGoster?: boolean;
}) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table>
        <thead>
          <tr>
            <th>Ürün Kodu</th>
            {seriGoster && <th>Seri</th>}
            <th>Koli Adet</th>
            <th>Koli Metraj</th>
            <th>Fiyat (USD/mt)</th>
          </tr>
        </thead>
        <tbody>
          {items.map((f) => (
            <tr key={f.code}>
              <td style={{ fontWeight: 600 }}>{f.code}</td>
              {seriGoster && <td>{f.series}</td>}
              <td>{f.koliAdet}</td>
              <td>{fmt(f.koliMetraj)} MT</td>
              <td>${fmt(f.priceUSD)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function TechTable({
  items,
  kategoriGoster,
}: {
  items: typeof TECHNICAL_PRODUCTS;
  kategoriGoster?: boolean;
}) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table>
        <thead>
          <tr>
            <th>Ürün</th>
            {kategoriGoster && <th>Marka / Kategori</th>}
            <th>Adet / Kutu</th>
            <th>Fiyat</th>
          </tr>
        </thead>
        <tbody>
          {items.map((t) => (
            <tr key={t.code}>
              <td style={{ fontWeight: 600 }}>{t.name}</td>
              {kategoriGoster && <td>{t.category}</td>}
              <td>{t.adetPerKutu.toLocaleString("tr-TR")}</td>
              <td>{techPrice(t)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function PriceListBrowser() {
  const [tab, setTab] = useState<"cerceve" | "teknik">("cerceve");
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
        ? TECHNICAL_PRODUCTS
        : TECHNICAL_PRODUCTS.filter((t) =>
            eslesir(query, t.name, t.code, t.category)
          ),
    [query, araniyor]
  );

  const toplam = frames.length + technicals.length;

  return (
    <>
      <div className="card no-print" style={{ marginBottom: 20 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Ara — çerçeve kodu (2315S, GB022) veya malzeme adı (askı teli, agraf)…"
          aria-label="Fiyat listesinde ara"
        />

        <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
          <button
            className={`btn small ${tab === "cerceve" ? "" : "secondary"}`}
            onClick={() => setTab("cerceve")}
          >
            Çerçeve Profilleri ({FRAME_PROFILES.length})
          </button>
          <button
            className={`btn small ${tab === "teknik" ? "" : "secondary"}`}
            onClick={() => setTab("teknik")}
          >
            🔧 Teknik Malzemeler ({TECHNICAL_PRODUCTS.length})
          </button>
        </div>

        <p style={{ margin: "10px 0 0", color: "var(--muted)", fontSize: 13 }}>
          {araniyor ? (
            toplam ? (
              <>
                <strong>{frames.length}</strong> çerçeve profili,{" "}
                <strong>{technicals.length}</strong> teknik malzeme bulundu.
                Arama iki listeyi birden tarar.
              </>
            ) : (
              <>Sonuç bulunamadı — farklı bir kod veya isim deneyin.</>
            )
          ) : tab === "cerceve" ? (
            <>Çerçeve profilleri seri seri listeleniyor. Fiyatlar USD/mt.</>
          ) : (
            <>
              Teknik malzemeler markaya göre listeleniyor. € Euro, ₺ Türk Lirası;
              fiyatlar kutu bazındadır.
            </>
          )}
        </p>
      </div>

      {/* ---------- Arama sonuçları: iki grup birlikte ---------- */}
      {araniyor && (
        <>
          {technicals.length > 0 && (
            <div className="card" style={{ marginBottom: 20 }}>
              <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>
                Teknik Malzemeler ({technicals.length})
              </h2>
              <TechTable items={technicals} kategoriGoster />
            </div>
          )}
          {frames.length > 0 && (
            <div className="card" style={{ marginBottom: 20 }}>
              <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>
                Çerçeve Profilleri ({frames.length})
              </h2>
              <FrameTable items={frames} seriGoster />
            </div>
          )}
        </>
      )}

      {/* ---------- Gezinme: seçili sekmenin tam listesi ---------- */}
      {!araniyor &&
        tab === "cerceve" &&
        SERIES_ORDER.map((series) => {
          const items = FRAME_PROFILES.filter((f) => f.series === series);
          if (!items.length) return null;
          return (
            <div className="card" key={series} style={{ marginBottom: 20 }}>
              <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>
                {series} Serisi
              </h2>
              <FrameTable items={items} />
            </div>
          );
        })}

      {!araniyor &&
        tab === "teknik" &&
        CATEGORY_ORDER.map((cat) => {
          const items = TECHNICAL_PRODUCTS.filter((t) => t.category === cat);
          if (!items.length) return null;
          return (
            <div className="card" key={cat} style={{ marginBottom: 20 }}>
              <h2 style={{ marginTop: 0, color: "var(--brand-light)" }}>{cat}</h2>
              <TechTable items={items} />
            </div>
          );
        })}
    </>
  );
}
