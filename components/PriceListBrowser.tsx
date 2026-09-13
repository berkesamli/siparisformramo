"use client";

// Toptan fiyat listesi — arama kutusunun altındaki iki kutu ile çerçeve
// profilleri ve teknik malzemeler arasında geçiş yapılır. Aynı sayfada kalınır;
// teknik malzemelere ulaşmak için aşağı kaydırmak gerekmez.
//
// Arama yapıldığında sekme farkı gözetilmez: her iki gruptan eşleşen ürünler
// birlikte listelenir, kullanıcı hangi sekmede olduğunu düşünmek zorunda kalmaz.

import { useEffect, useMemo, useState, type CSSProperties, type ReactNode } from "react";
// Fiyat listesi istemci paketinde durmaz — oturumla /api/katalog'dan gelir.
import {
  SERIES_ORDER,
  type FrameProfile,
  type TechnicalProduct,
} from "@/lib/catalog-utils";
import { useKatalog } from "@/lib/use-katalog";
import { eslesir } from "@/lib/search-norm";
import Icon from "@/components/shell/Icon";

const fmt = (n: number) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const techPrice = (t: TechnicalProduct): string =>
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

type Sekme = "cerceve" | "teknik";

// Dış dünyanın "profil" | "teknik" değerini iç sekme değerine çevirir;
// "teknik" dışındaki her şey çerçeve profilleridir.
const sekmeyeCevir = (l?: string): Sekme => (l === "teknik" ? "teknik" : "cerceve");

function FrameTable({
  items,
  seriGoster,
}: {
  items: FrameProfile[];
  seriGoster?: boolean;
}) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Ürün Kodu</th>
            {seriGoster && <th>Seri</th>}
            <th className="num">Koli Adet</th>
            <th className="num">Koli Metraj</th>
            <th className="num">Fiyat (USD/mt)</th>
          </tr>
        </thead>
        <tbody>
          {items.map((f) => (
            <tr key={f.code}>
              <td style={{ fontWeight: 600 }}>{f.code}</td>
              {seriGoster && <td>{f.series}</td>}
              <td className="num">{f.koliAdet}</td>
              <td className="num">{fmt(f.koliMetraj)} MT</td>
              <td className="num">${fmt(f.priceUSD)}</td>
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
  items: TechnicalProduct[];
  kategoriGoster?: boolean;
}) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Ürün</th>
            {kategoriGoster && <th>Marka / Kategori</th>}
            <th className="num">Adet / Kutu</th>
            <th className="num">Fiyat</th>
          </tr>
        </thead>
        <tbody>
          {items.map((t) => (
            <tr key={t.code}>
              <td style={{ fontWeight: 600 }}>{t.name}</td>
              {kategoriGoster && <td>{t.category}</td>}
              <td className="num">{t.adetPerKutu.toLocaleString("tr-TR")}</td>
              <td className="num">{techPrice(t)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function ListeBaslik({
  icon,
  children,
}: {
  icon: "frame" | "settings";
  children: ReactNode;
}) {
  return (
    <div className="card-head">
      <span className="card-head-icon">
        <Icon name={icon} size={16} />
      </span>
      <div>
        <h2>{children}</h2>
      </div>
    </div>
  );
}

export default function PriceListBrowser({
  initialQuery,
  initialList,
}: {
  /** Arama kutusunu tohumlar (sunucu sayfası ?q= parametresinden geçirir). */
  initialQuery?: string;
  /** Başlangıç listesi: "teknik" → teknik malzemeler, diğerleri → çerçeve profilleri. */
  initialList?: "profil" | "teknik";
}) {
  const katalog = useKatalog();
  const [tab, setTab] = useState<Sekme>(() => sekmeyeCevir(initialList));
  const [query, setQuery] = useState(initialQuery || "");
  const araniyor = query.trim().length > 0;

  // URL parametreleri değişirse (üst çubuk aramasından yeniden gelinirse) eşitle.
  useEffect(() => {
    if (initialQuery !== undefined) setQuery(initialQuery);
  }, [initialQuery]);
  useEffect(() => {
    if (initialList !== undefined) setTab(sekmeyeCevir(initialList));
  }, [initialList]);

  const CATEGORY_ORDER = useMemo(
    () => [
      ...BILINEN_SIRA,
      ...Array.from(new Set(katalog.technical.map((t) => t.category))).filter(
        (c) => !BILINEN_SIRA.includes(c)
      ),
    ],
    [katalog.technical]
  );

  const frames = useMemo(
    () =>
      !araniyor
        ? katalog.profiles
        : katalog.profiles.filter((f) => eslesir(query, f.code, f.series)),
    [query, araniyor, katalog.profiles]
  );

  const technicals = useMemo(
    () =>
      !araniyor
        ? katalog.technical
        : katalog.technical.filter((t) =>
            eslesir(query, t.name, t.code, t.category)
          ),
    [query, araniyor, katalog.technical]
  );

  const toplam = frames.length + technicals.length;

  if (!katalog.yuklendi) {
    return (
      <div className="card">
        <p style={{ color: "var(--muted)", margin: 0 }}>Fiyat listesi yükleniyor…</p>
      </div>
    );
  }

  const segBtn: CSSProperties = {
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
  };

  return (
    <>
      <div className="card no-print">
        <div style={{ position: "relative" }}>
          <span
            aria-hidden
            style={{
              position: "absolute",
              left: 12,
              top: "50%",
              transform: "translateY(-50%)",
              color: "var(--muted)",
              display: "inline-flex",
              pointerEvents: "none",
            }}
          >
            <Icon name="search" size={16} />
          </span>
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Ara — çerçeve kodu (2315S, GB022) veya malzeme adı (askı teli, agraf)…"
            aria-label="Fiyat listesinde ara"
            style={{ paddingLeft: 38 }}
          />
        </div>

        <div className="row" style={{ marginTop: 12 }}>
          <div className="seg" role="tablist" aria-label="Fiyat listesi">
            <button
              type="button"
              role="tab"
              aria-selected={tab === "cerceve"}
              className={tab === "cerceve" ? "active" : ""}
              style={segBtn}
              onClick={() => setTab("cerceve")}
            >
              <Icon name="frame" size={14} />
              Çerçeve Profilleri ({katalog.profiles.length})
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={tab === "teknik"}
              className={tab === "teknik" ? "active" : ""}
              style={segBtn}
              onClick={() => setTab("teknik")}
            >
              <Icon name="settings" size={14} />
              Teknik Malzemeler ({katalog.technical.length})
            </button>
          </div>
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
            <div className="card">
              <ListeBaslik icon="settings">
                Teknik Malzemeler ({technicals.length})
              </ListeBaslik>
              <TechTable items={technicals} kategoriGoster />
            </div>
          )}
          {frames.length > 0 && (
            <div className="card">
              <ListeBaslik icon="frame">
                Çerçeve Profilleri ({frames.length})
              </ListeBaslik>
              <FrameTable items={frames} seriGoster />
            </div>
          )}
        </>
      )}

      {/* ---------- Gezinme: seçili sekmenin tam listesi ---------- */}
      {!araniyor &&
        tab === "cerceve" &&
        SERIES_ORDER.map((series) => {
          const items = katalog.profiles.filter((f) => f.series === series);
          if (!items.length) return null;
          return (
            <div className="card" key={series}>
              <ListeBaslik icon="frame">{series} Serisi</ListeBaslik>
              <FrameTable items={items} />
            </div>
          );
        })}

      {!araniyor &&
        tab === "teknik" &&
        CATEGORY_ORDER.map((cat) => {
          const items = katalog.technical.filter((t) => t.category === cat);
          if (!items.length) return null;
          return (
            <div className="card" key={cat}>
              <ListeBaslik icon="settings">{cat}</ListeBaslik>
              <TechTable items={items} />
            </div>
          );
        })}
    </>
  );
}
