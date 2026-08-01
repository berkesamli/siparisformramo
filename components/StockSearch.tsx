"use client";

// Günlük stok sorgulama — Ankara/İstanbul depoları, metraj 2,9'a bölünüp
// "boy" olarak gösterilir. Bulanık arama: "gc065-1473" → GC065-1473BX.

import { useEffect, useMemo, useState } from "react";
import type { StockData } from "@/lib/stock-parse";
import { searchStock, toBoy } from "@/lib/stock-search";

const nf = (n: number) => n.toLocaleString("tr-TR");

export default function StockSearch() {
  const [data, setData] = useState<StockData | null>(null);
  const [error, setError] = useState("");
  const [query, setQuery] = useState("");

  useEffect(() => {
    fetch("/api/stock")
      .then((r) => r.json())
      .then((d) => {
        if (d.ok) setData(d.data);
        else setError(d.error || "Stok verisi alınamadı.");
      })
      .catch(() => setError("Stok verisi alınamadı."));
  }, []);

  const results = useMemo(() => {
    if (!data) return [];
    return searchStock(data.items, query);
  }, [data, query]);

  const updatedStr = data
    ? new Date(data.updatedAt).toLocaleString("tr-TR", {
        dateStyle: "medium",
        timeStyle: "short",
        timeZone: "Europe/Istanbul",
      })
    : "";

  return (
    <div>
      <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 14 }}>
        <input
          style={{ maxWidth: 340 }}
          placeholder="Profil kodu yazın… örn. GC065-1473"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          autoFocus
        />
        {data && (
          <span style={{ color: "var(--muted)", fontSize: 12.5 }}>
            {nf(data.items.length)} kalem · Son güncelleme: {updatedStr}
          </span>
        )}
      </div>

      {error && <div className="notice err">{error}</div>}
      {!data && !error && <p style={{ color: "var(--text-2)" }}>Stok verisi yükleniyor…</p>}

      {data && query.trim().length >= 2 && (
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Ürün Kodu</th>
                <th>Ankara Depo</th>
                <th>İstanbul Depo</th>
                <th>Toplam</th>
              </tr>
            </thead>
            <tbody>
              {results.map(({ item, score }) => {
                const total = item.ankaraMt + item.istanbulMt;
                return (
                  <tr key={item.code}>
                    <td style={{ fontWeight: 600 }}>
                      {item.code}
                      {score < 0.9 && (
                        <span style={{ color: "var(--muted)", fontSize: 11.5, marginLeft: 6 }}>
                          (benzer)
                        </span>
                      )}
                    </td>
                    <td>
                      {item.ankaraMt > 0 ? (
                        <><strong>{nf(toBoy(item.ankaraMt))} boy</strong></>
                      ) : (
                        <span className="badge yok">Yok</span>
                      )}
                    </td>
                    <td>
                      {item.istanbulMt > 0 ? (
                        <><strong>{nf(toBoy(item.istanbulMt))} boy</strong></>
                      ) : (
                        <span className="badge yok">Yok</span>
                      )}
                    </td>
                    <td>
                      <strong style={{ color: "var(--brand-light)" }}>
                        {nf(toBoy(total))} boy
                      </strong>
                    </td>
                  </tr>
                );
              })}
              {!results.length && (
                <tr>
                  <td colSpan={4} style={{ color: "var(--muted)" }}>
                    &quot;{query}&quot; ile eşleşen profil bulunamadı. Güncel bilgi için:
                    0850 305 75 45
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {data && query.trim().length < 2 && (
        <p style={{ color: "var(--muted)", fontSize: 13 }}>
          Aramak için profil kodunu yazmaya başlayın — kodu eksik ya da hatalı
          yazsanız bile en yakın eşleşmeleri gösteririz.
        </p>
      )}
    </div>
  );
}
