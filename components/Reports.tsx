"use client";

// Rapor ekranı: ciro özeti, aylık dağılım, müşteri/ürün/seri/çalışan kırılımları.

import { useCallback, useEffect, useMemo, useState } from "react";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const k = (n: number) =>
  n >= 1000
    ? (n / 1000).toLocaleString("tr-TR", { maximumFractionDigits: 1 }) + "b"
    : String(Math.round(n));

interface Data {
  blob: boolean;
  summary: {
    orderCount: number;
    toptanCount: number;
    perakendeCount: number;
    toptanCiro: number;
    perakendeCiro: number;
    toplamCiro: number;
    tahsilat: number;
    bakiye: number;
    ortalamaSepet: number;
  };
  months: { month: string; toptan: number; perakende: number; toplam: number }[];
  customers: { name: string; total: number; count: number; balance: number }[];
  products: { name: string; total: number; count: number }[];
  series: { name: string; total: number }[];
  employees: { name: string; total: number; count: number }[];
}

const AY_ADI = [
  "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
  "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık",
];
const ayLabel = (m: string) => {
  const [y, mo] = m.split("-");
  return `${AY_ADI[Number(mo) - 1] || mo} ${y}`;
};

export default function Reports() {
  const [ay, setAy] = useState(""); // "" = tüm zamanlar
  const [data, setData] = useState<Data | null>(null);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch(`/api/raporlar${ay ? `?ay=${ay}` : ""}`);
      if (res.ok) setData(await res.json());
    } finally {
      setLoading(false);
    }
  }, [ay]);

  useEffect(() => {
    load();
  }, [load]);

  const maxMonth = useMemo(
    () => Math.max(1, ...(data?.months || []).map((m) => m.toplam)),
    [data]
  );

  const thisMonth = new Date().toISOString().slice(0, 7);

  return (
    <div>
      {/* Dönem seçimi */}
      <div className="card" style={{ padding: 14, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
        <button className={`btn small ${ay === "" ? "" : "secondary"}`} onClick={() => setAy("")}>
          Tüm Zamanlar
        </button>
        <button
          className={`btn small ${ay === thisMonth ? "" : "secondary"}`}
          onClick={() => setAy(thisMonth)}
        >
          Bu Ay
        </button>
        <input
          type="month"
          style={{ width: 180 }}
          value={ay}
          onChange={(e) => setAy(e.target.value)}
        />
        <span style={{ flex: 1 }} />
        <button className="btn small secondary" onClick={load}>↻ Yenile</button>
      </div>

      {loading && <p style={{ color: "var(--muted)" }}>Hesaplanıyor...</p>}

      {data && !loading && (
        <>
          {!data.blob && (
            <div className="notice info">
              Kalıcı depolama yapılandırılmadığı için rapor verisi okunamıyor.
            </div>
          )}

          {/* Özet kutuları */}
          <div className="cari-cards">
            <div className="cari-card">
              <span>Toplam Ciro</span>
              <strong style={{ color: "var(--brand)" }}>₺{fmt(data.summary.toplamCiro)}</strong>
            </div>
            <div className="cari-card">
              <span>Sipariş Sayısı</span>
              <strong>{data.summary.orderCount}</strong>
            </div>
            <div className="cari-card">
              <span>Tahsil Edilen</span>
              <strong style={{ color: "var(--success)" }}>₺{fmt(data.summary.tahsilat)}</strong>
            </div>
            <div className={`cari-card ${data.summary.bakiye > 0 ? "borc" : ""}`}>
              <span>Açık Bakiye</span>
              <strong style={{ color: data.summary.bakiye > 0 ? "var(--error)" : "var(--success)" }}>
                ₺{fmt(data.summary.bakiye)}
              </strong>
            </div>
            <div className="cari-card">
              <span>Ortalama Sipariş</span>
              <strong>₺{fmt(data.summary.ortalamaSepet)}</strong>
            </div>
          </div>

          {/* Toptan / perakende kırılımı */}
          <div className="rep-split">
            <div className="rep-split-item">
              <div className="rep-split-head">
                <b>Toptan</b>
                <span>{data.summary.toptanCount} sipariş</span>
              </div>
              <strong>₺{fmt(data.summary.toptanCiro)}</strong>
              <div className="rep-bar">
                <span
                  className="seg toptan"
                  style={{
                    width: `${data.summary.toplamCiro ? (data.summary.toptanCiro / data.summary.toplamCiro) * 100 : 0}%`,
                  }}
                />
              </div>
            </div>
            <div className="rep-split-item">
              <div className="rep-split-head">
                <b>Perakende</b>
                <span>{data.summary.perakendeCount} sipariş</span>
              </div>
              <strong>₺{fmt(data.summary.perakendeCiro)}</strong>
              <div className="rep-bar">
                <span
                  className="seg perakende"
                  style={{
                    width: `${data.summary.toplamCiro ? (data.summary.perakendeCiro / data.summary.toplamCiro) * 100 : 0}%`,
                  }}
                />
              </div>
            </div>
          </div>

          {/* Aylık ciro grafiği */}
          <h2>Aylık Ciro</h2>
          {data.months.length === 0 ? (
            <div className="card" style={{ color: "var(--muted)", textAlign: "center" }}>
              Henüz sipariş kaydı yok.
            </div>
          ) : (
            <div className="card">
              <div className="rep-chart">
                {[...data.months].reverse().map((m) => (
                  <div className="rep-col" key={m.month} title={`${ayLabel(m.month)}: ₺${fmt(m.toplam)}`}>
                    <div className="rep-col-bars">
                      <span
                        className="seg toptan"
                        style={{ height: `${(m.toptan / maxMonth) * 100}%` }}
                      />
                      <span
                        className="seg perakende"
                        style={{ height: `${(m.perakende / maxMonth) * 100}%` }}
                      />
                    </div>
                    <div className="rep-col-val">₺{k(m.toplam)}</div>
                    <div className="rep-col-lbl">{ayLabel(m.month).slice(0, 3)}</div>
                  </div>
                ))}
              </div>
              <div className="rep-legend">
                <span><i className="dot toptan" /> Toptan</span>
                <span><i className="dot perakende" /> Perakende</span>
              </div>
            </div>
          )}

          {/* Tablolar */}
          <div className="rep-tables">
            <div className="card" style={{ padding: 0 }}>
              <h3 className="rep-th">En Çok Alan Müşteriler</h3>
              <table>
                <thead>
                  <tr><th>Müşteri</th><th style={{ textAlign: "right" }}>Sipariş</th><th style={{ textAlign: "right" }}>Ciro</th><th style={{ textAlign: "right" }}>Bakiye</th></tr>
                </thead>
                <tbody>
                  {data.customers.length === 0 ? (
                    <tr><td colSpan={4} style={{ color: "var(--muted)" }}>Kayıt yok</td></tr>
                  ) : (
                    data.customers.map((c) => (
                      <tr key={c.name}>
                        <td style={{ fontWeight: 600 }}>{c.name}</td>
                        <td style={{ textAlign: "right" }}>{c.count}</td>
                        <td style={{ textAlign: "right" }}>₺{fmt(c.total)}</td>
                        <td style={{ textAlign: "right", color: c.balance > 0 ? "var(--error)" : "var(--muted)" }}>
                          {c.balance > 0 ? `₺${fmt(c.balance)}` : "—"}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            <div className="card" style={{ padding: 0 }}>
              <h3 className="rep-th">En Çok Satan Ürünler</h3>
              <table>
                <thead>
                  <tr><th>Ürün</th><th style={{ textAlign: "right" }}>Satır</th><th style={{ textAlign: "right" }}>Tutar</th></tr>
                </thead>
                <tbody>
                  {data.products.length === 0 ? (
                    <tr><td colSpan={3} style={{ color: "var(--muted)" }}>Kayıt yok</td></tr>
                  ) : (
                    data.products.map((p) => (
                      <tr key={p.name}>
                        <td>{p.name}</td>
                        <td style={{ textAlign: "right" }}>{p.count}</td>
                        <td style={{ textAlign: "right" }}>₺{fmt(p.total)}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {data.series.length > 0 && (
              <div className="card" style={{ padding: 0 }}>
                <h3 className="rep-th">Seri Bazlı Satış</h3>
                <table>
                  <thead>
                    <tr><th>Seri</th><th style={{ textAlign: "right" }}>Tutar</th></tr>
                  </thead>
                  <tbody>
                    {data.series.map((s) => (
                      <tr key={s.name}>
                        <td style={{ fontWeight: 600 }}>{s.name} Serisi</td>
                        <td style={{ textAlign: "right" }}>₺{fmt(s.total)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            <div className="card" style={{ padding: 0 }}>
              <h3 className="rep-th">Çalışan Performansı</h3>
              <table>
                <thead>
                  <tr><th>Çalışan</th><th style={{ textAlign: "right" }}>Sipariş</th><th style={{ textAlign: "right" }}>Ciro</th></tr>
                </thead>
                <tbody>
                  {data.employees.length === 0 ? (
                    <tr><td colSpan={3} style={{ color: "var(--muted)" }}>Kayıt yok</td></tr>
                  ) : (
                    data.employees.map((e) => (
                      <tr key={e.name}>
                        <td style={{ fontWeight: 600 }}>{e.name}</td>
                        <td style={{ textAlign: "right" }}>{e.count}</td>
                        <td style={{ textAlign: "right" }}>₺{fmt(e.total)}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}
