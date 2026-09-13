"use client";

// Rapor ekranı: ciro özeti, aylık dağılım, müşteri/ürün/seri/çalışan kırılımları.

import { useCallback, useEffect, useMemo, useState } from "react";
import Icon from "@/components/shell/Icon";

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
    gercekTahsilat: number;
    giderToplam: number;
    kasaKar: number;
    faturaliCiro: number;
    faturasizCiro: number;
    bakiye: number;
    ortalamaSepet: number;
  };
  months: { month: string; toptan: number; perakende: number; toplam: number }[];
  customers: { name: string; total: number; count: number; balance: number }[];
  products: { name: string; total: number; count: number }[];
  series: { name: string; total: number }[];
  employees: { name: string; total: number; count: number }[];
  tahsilEdenler: { name: string; total: number; count: number; yontem: Record<string, number> }[];
}

const AY_ADI = [
  "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
  "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık",
];
const ayLabel = (m: string) => {
  const [y, mo] = m.split("-");
  return `${AY_ADI[Number(mo) - 1] || mo} ${y}`;
};

// Telefonda KPI kutuları iki sütuna sığsın: orders.css'teki .cari-cards
// eşiği (180px) yerine 160px; rakam boyutu dar ekranda akışkan küçülür
// (390px'te ~17px, 480px'ten itibaren 21px) ki uzun tutarlar bölünmesin.
const KPI_GRID = { gridTemplateColumns: "repeat(auto-fit, minmax(min(160px, 100%), 1fr))" } as const;
const KPI_NUM = { fontSize: "clamp(15px, 4.4vw, 21px)" } as const;
// Sıfır dolgulu kart içinde .table-wrap'in taşma payı (−8px) gereksiz.
const WRAP = { margin: 0, padding: 0 } as const;
const NOWRAP = { whiteSpace: "nowrap" } as const;

export default function Reports() {
  const [ay, setAy] = useState(""); // "" = tüm zamanlar
  const [sube, setSube] = useState("");
  const [data, setData] = useState<Data | null>(null);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (ay) params.set("ay", ay);
      if (sube) params.set("sube", sube);
      const qs = params.toString();
      const res = await fetch(`/api/raporlar${qs ? `?${qs}` : ""}`);
      if (res.ok) setData(await res.json());
    } finally {
      setLoading(false);
    }
  }, [ay, sube]);

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
      <div className="card pad-sm row no-print">
        <div className="seg">
          <button className={ay === "" ? "active" : undefined} onClick={() => setAy("")}>
            Tüm Zamanlar
          </button>
          <button className={ay === thisMonth ? "active" : undefined} onClick={() => setAy(thisMonth)}>
            Bu Ay
          </button>
        </div>
        <input
          type="month"
          style={{ width: "auto", maxWidth: "100%" }}
          value={ay}
          onChange={(e) => setAy(e.target.value)}
        />
        <select style={{ width: "auto", maxWidth: "100%" }} value={sube} onChange={(e) => setSube(e.target.value)}>
          <option value="">Tüm Şubeler</option>
          <option value="ankara">Ankara</option>
          <option value="istanbul">İstanbul</option>
          <option value="belirsiz">Şubesiz (eski)</option>
        </select>
        <button className="btn small secondary" style={{ marginLeft: "auto" }} onClick={load}>
          <Icon name="refresh" size={14} /> Yenile
        </button>
      </div>

      {loading && <p className="muted">Hesaplanıyor...</p>}

      {data && !loading && (
        <>
          {!data.blob && (
            <div className="notice info">
              Kalıcı depolama yapılandırılmadığı için rapor verisi okunamıyor.
            </div>
          )}

          {/* Özet kutuları */}
          <div className="cari-cards" style={KPI_GRID}>
            <div className="cari-card">
              <span>Toplam Ciro</span>
              <strong style={{ ...KPI_NUM, color: "var(--brand)" }}>₺{fmt(data.summary.toplamCiro)}</strong>
            </div>
            <div className="cari-card">
              <span>Sipariş Sayısı</span>
              <strong style={KPI_NUM}>{data.summary.orderCount}</strong>
            </div>
            <div className="cari-card">
              <span>Tahsil Edilen</span>
              <strong style={{ ...KPI_NUM, color: "var(--success)" }}>₺{fmt(data.summary.tahsilat)}</strong>
            </div>
            <div className={`cari-card ${data.summary.bakiye > 0 ? "borc" : ""}`}>
              <span>Açık Bakiye</span>
              <strong style={{ ...KPI_NUM, color: data.summary.bakiye > 0 ? "var(--error)" : "var(--success)" }}>
                ₺{fmt(data.summary.bakiye)}
              </strong>
            </div>
            <div className="cari-card">
              <span>Ortalama Sipariş</span>
              <strong style={KPI_NUM}>₺{fmt(data.summary.ortalamaSepet)}</strong>
            </div>
          </div>

          {/* Kasa bazlı satır — gerçek tahsilat/gider kayıtlarından */}
          <div className="cari-cards" style={KPI_GRID}>
            <div className="cari-card">
              <span>Kasa Tahsilatı</span>
              <strong style={{ ...KPI_NUM, color: "var(--success)" }}>₺{fmt(data.summary.gercekTahsilat || 0)}</strong>
              <span style={{ fontSize: 11.5 }}>tahsilat kayıtlarından</span>
            </div>
            <div className="cari-card">
              <span>Giderler</span>
              <strong style={{ ...KPI_NUM, color: "var(--error)" }}>₺{fmt(data.summary.giderToplam || 0)}</strong>
              {!ay && <span style={{ fontSize: 11.5 }}>son 12 ay</span>}
            </div>
            <div className={`cari-card ${(data.summary.kasaKar || 0) < 0 ? "borc" : ""}`}>
              <span>Kasa Kârı</span>
              <strong style={{ ...KPI_NUM, color: (data.summary.kasaKar || 0) >= 0 ? "var(--success)" : "var(--error)" }}>
                ₺{fmt(data.summary.kasaKar || 0)}
              </strong>
              <span style={{ fontSize: 11.5 }}>tahsilat − gider</span>
            </div>
            <div className="cari-card">
              <span>Faturalı Ciro</span>
              <strong style={KPI_NUM}>₺{fmt(data.summary.faturaliCiro || 0)}</strong>
              <span style={{ fontSize: 11.5 }}>KDV&apos;li siparişler</span>
            </div>
            <div className="cari-card">
              <span>Faturasız Ciro</span>
              <strong style={KPI_NUM}>₺{fmt(data.summary.faturasizCiro || 0)}</strong>
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
          {data.months.length === 0 ? (
            <div className="card">
              <div className="empty">
                <div className="empty-icon"><Icon name="bar-chart" size={24} /></div>
                <strong>Aylık Ciro</strong>
                Henüz sipariş kaydı yok.
              </div>
            </div>
          ) : (
            <div className="card">
              <div className="card-head">
                <span className="card-head-icon"><Icon name="bar-chart" size={18} /></span>
                <div><h2>Aylık Ciro</h2></div>
              </div>
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
            <div className="card pad-0">
              <h3 className="rep-th">En Çok Alan Müşteriler</h3>
              <div className="table-wrap" style={WRAP}>
                <table>
                  <thead>
                    <tr><th>Müşteri</th><th className="num">Sipariş</th><th className="num">Ciro</th><th className="num">Bakiye</th></tr>
                  </thead>
                  <tbody>
                    {data.customers.length === 0 ? (
                      <tr><td colSpan={4} className="muted">Kayıt yok</td></tr>
                    ) : (
                      data.customers.map((c) => (
                        <tr key={c.name}>
                          <td style={{ fontWeight: 600 }}>{c.name}</td>
                          <td className="num">{c.count}</td>
                          <td className="num" style={NOWRAP}>₺{fmt(c.total)}</td>
                          <td className="num" style={{ ...NOWRAP, color: c.balance > 0 ? "var(--error)" : "var(--muted)" }}>
                            {c.balance > 0 ? `₺${fmt(c.balance)}` : "—"}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="card pad-0">
              <h3 className="rep-th">En Çok Satan Ürünler</h3>
              <div className="table-wrap" style={WRAP}>
                <table>
                  <thead>
                    <tr><th>Ürün</th><th className="num">Satır</th><th className="num">Tutar</th></tr>
                  </thead>
                  <tbody>
                    {data.products.length === 0 ? (
                      <tr><td colSpan={3} className="muted">Kayıt yok</td></tr>
                    ) : (
                      data.products.map((p) => (
                        <tr key={p.name}>
                          <td>{p.name}</td>
                          <td className="num">{p.count}</td>
                          <td className="num" style={NOWRAP}>₺{fmt(p.total)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {data.series.length > 0 && (
              <div className="card pad-0">
                <h3 className="rep-th">Seri Bazlı Satış</h3>
                <div className="table-wrap" style={WRAP}>
                  <table>
                    <thead>
                      <tr><th>Seri</th><th className="num">Tutar</th></tr>
                    </thead>
                    <tbody>
                      {data.series.map((s) => (
                        <tr key={s.name}>
                          <td style={{ fontWeight: 600 }}>{s.name} Serisi</td>
                          <td className="num" style={NOWRAP}>₺{fmt(s.total)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="card pad-0">
              <h3 className="rep-th">Çalışan Performansı</h3>
              <div className="table-wrap" style={WRAP}>
                <table>
                  <thead>
                    <tr><th>Çalışan</th><th className="num">Sipariş</th><th className="num">Ciro</th></tr>
                  </thead>
                  <tbody>
                    {data.employees.length === 0 ? (
                      <tr><td colSpan={3} className="muted">Kayıt yok</td></tr>
                    ) : (
                      data.employees.map((e) => (
                        <tr key={e.name}>
                          <td style={{ fontWeight: 600 }}>{e.name}</td>
                          <td className="num">{e.count}</td>
                          <td className="num" style={NOWRAP}>₺{fmt(e.total)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Prim raporu — Alaattin'in prim çizelgesinin karşılığı */}
            <div className="card pad-0">
              <h3 className="rep-th">Tahsil Eden Bazlı (Prim)</h3>
              <div className="table-wrap" style={WRAP}>
                <table>
                  <thead>
                    <tr><th>Tahsil Eden</th><th className="num">İşlem</th><th className="num">Tahsilat</th></tr>
                  </thead>
                  <tbody>
                    {(data.tahsilEdenler || []).length === 0 ? (
                      <tr><td colSpan={3} className="muted">Kayıt yok</td></tr>
                    ) : (
                      (data.tahsilEdenler || []).map((e) => (
                        <tr key={e.name}>
                          <td style={{ fontWeight: 600 }}>{e.name}</td>
                          <td className="num">{e.count}</td>
                          <td className="num" style={NOWRAP}>₺{fmt(e.total)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </>
      )}
    </div>
  );
}
