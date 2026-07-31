"use client";

import { useMemo, useState } from "react";
import { FRAME_PROFILES, SERIES_ORDER, boyLength } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";

const fmt = (n: number, d = 2) =>
  n.toLocaleString("tr-TR", { minimumFractionDigits: d, maximumFractionDigits: d });

const STOCK_LABEL: Record<string, string> = {
  var: "Stokta",
  az: "Az Kaldı",
  yok: "Tükendi",
};

export default function ProductBrowser() {
  const [tab, setTab] = useState<"frame" | "technical">("frame");
  const [series, setSeries] = useState("all");
  const [query, setQuery] = useState("");

  const frames = useMemo(() => {
    const q = query.trim().toLowerCase();
    return FRAME_PROFILES.filter(
      (f) =>
        (series === "all" || f.series === series) &&
        (!q || f.code.toLowerCase().includes(q))
    );
  }, [series, query]);

  const technicals = useMemo(() => {
    const q = query.trim().toLowerCase();
    return TECHNICAL_PRODUCTS.filter(
      (t) =>
        !q ||
        t.name.toLowerCase().includes(q) ||
        t.category.toLowerCase().includes(q)
    );
  }, [query]);

  return (
    <div className="card">
      <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <button
          className={`btn small ${tab === "frame" ? "" : "secondary"}`}
          onClick={() => setTab("frame")}
        >
          Çerçeve Profilleri ({FRAME_PROFILES.length})
        </button>
        <button
          className={`btn small ${tab === "technical" ? "" : "secondary"}`}
          onClick={() => setTab("technical")}
        >
          Teknik Malzemeler ({TECHNICAL_PRODUCTS.length})
        </button>
        <span style={{ flex: 1 }} />
        {tab === "frame" && (
          <select
            style={{ width: "auto" }}
            value={series}
            onChange={(e) => setSeries(e.target.value)}
          >
            <option value="all">Tüm Seriler</option>
            {SERIES_ORDER.map((s) => (
              <option key={s} value={s}>
                {s} Serisi
              </option>
            ))}
          </select>
        )}
        <input
          style={{ width: 220 }}
          placeholder="Ara…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
      </div>

      {tab === "frame" ? (
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Ürün Kodu</th>
                <th>Seri</th>
                <th>Koli Adet</th>
                <th>Koli Metraj</th>
                <th>Boy Uzunluğu</th>
                <th>Toptan Fiyat</th>
                <th>Stok</th>
              </tr>
            </thead>
            <tbody>
              {frames.map((f) => (
                <tr key={f.code}>
                  <td style={{ fontWeight: 600 }}>{f.code}</td>
                  <td>{f.series}</td>
                  <td>{f.koliAdet}</td>
                  <td>{fmt(f.koliMetraj)} mt</td>
                  <td>{fmt(boyLength(f))} mt</td>
                  <td>${fmt(f.priceUSD)}/mt + KDV</td>
                  <td>
                    <span className={`badge ${f.stok}`}>{STOCK_LABEL[f.stok]}</span>
                  </td>
                </tr>
              ))}
              {!frames.length && (
                <tr>
                  <td colSpan={7} style={{ color: "var(--muted)" }}>
                    Sonuç bulunamadı.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Ürün</th>
                <th>Marka / Kategori</th>
                <th>Adet / Kutu</th>
                <th>Fiyat</th>
                <th>Stok</th>
              </tr>
            </thead>
            <tbody>
              {technicals.map((t) => (
                <tr key={t.code}>
                  <td style={{ fontWeight: 600 }}>{t.name}</td>
                  <td>{t.category}</td>
                  <td>{t.adetPerKutu.toLocaleString("tr-TR")}</td>
                  <td>
                    {t.priceTL != null
                      ? `₺${fmt(t.priceTL)}`
                      : `€${fmt(t.priceEUR || 0)}`}
                  </td>
                  <td>
                    <span className={`badge ${t.stok || "var"}`}>
                      {STOCK_LABEL[t.stok || "var"]}
                    </span>
                  </td>
                </tr>
              ))}
              {!technicals.length && (
                <tr>
                  <td colSpan={5} style={{ color: "var(--muted)" }}>
                    Sonuç bulunamadı.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
