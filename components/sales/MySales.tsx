"use client";

// Satışlarım — çalışanın yalnızca KENDİ adına girilen siparişleri.
// Veri /api/cirom'dan gelir (ad oturumdan alınır; başka çalışan sorgulanamaz).
// Ay değişince önceki çizim solgun kalır (iskelet flaşı yok).

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import type { KisiselOzet, AyOzet } from "@/lib/personal-sales";
import StatTile from "@/components/dashboard/StatTile";
import OrderFlowChart from "@/components/dashboard/OrderFlowChart";
import { useSize } from "@/components/dashboard/useSize";

const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });
const tlOrt = (n: number) => {
  const v = Number(n) || 0;
  return v < 1000
    ? "₺" + v.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : tl(v);
};
const nf = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");
const saat = (iso: string) => {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? "" : d.toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit", timeZone: "Europe/Istanbul" });
};
const gunKisa = (k: string) => { const [, m, d] = k.split("-"); return `${Number(d)}.${m}`; };

const BOS_DONEM = { adet: 0, ciro: 0, toptanAdet: 0, perakendeAdet: 0, toptanCiro: 0, perakendeCiro: 0 };
const BOS: KisiselOzet = {
  employee: "", today: "", ay: "", ayLabel: "Bu Ay", aylar: [],
  bugun: BOS_DONEM, hafta: BOS_DONEM, secili: BOS_DONEM, oncekiAy: BOS_DONEM,
  ortalamaSiparis: 0, enIyiGun: null, seri14: [], sonSiparisler: [], blob: false, hesaplandi: "",
};

function niceMax(v: number): number {
  if (v <= 0) return 4;
  const p = Math.pow(10, Math.floor(Math.log10(v)));
  const m = v / p;
  const step = m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10;
  const top = step * p;
  return top === v ? v + (top / 4 || 1) : top;
}
const tickFmt = (v: number) => (v >= 1000000 ? `${(v / 1000000).toLocaleString("tr-TR", { maximumFractionDigits: 1 })}M` : v >= 1000 ? `${Math.round(v / 1000)}k` : String(Math.round(v)));

/** Yalnızca üst köşeleri yuvarlatılmış sütun (taban çizgisine oturur). */
function ustYuvarlak(x: number, y: number, w: number, h: number, r: number): string {
  const rr = Math.max(0, Math.min(r, w / 2, h));
  return `M${x},${y + h} V${y + rr} Q${x},${y} ${x + rr},${y} H${x + w - rr} Q${x + w},${y} ${x + w},${y + rr} V${y + h} Z`;
}

/* ---------- Aylık ciro sütun grafiği (tek seri) ---------- */
function MonthChart({
  aylar, secili, table, onSelect,
}: {
  aylar: AyOzet[];
  secili: string;
  table: boolean;
  onSelect: (ay: string) => void;
}) {
  const { ref, width } = useSize<HTMLDivElement>();
  const [hover, setHover] = useState<number | null>(null);

  const H = 230;
  const padL = 46, padR = 10, padT = 22, padB = 30;
  const plotW = Math.max(0, width - padL - padR);
  const plotH = H - padT - padB;
  const n = aylar.length || 1;
  const band = plotW / n;
  const barW = Math.min(64, Math.max(10, band * 0.56));
  const maxV = niceMax(Math.max(0, ...aylar.map((a) => a.ciro)));
  const ticks = [0, 0.25, 0.5, 0.75, 1].map((t) => t * maxV);
  const y = (v: number) => padT + plotH - (v / maxV) * plotH;
  const enYuksek = aylar.reduce((m, a) => (a.ciro > m ? a.ciro : m), 0);
  const tip = hover !== null ? aylar[hover] : null;

  return (
    <>
      {table && (
        <div className="table-wrap">
          <table className="chart-table mc-table">
            <thead>
              <tr><th>Ay</th><th className="num">Sipariş</th><th className="num">Toptan</th><th className="num">Perakende</th><th className="num">Toplam</th></tr>
            </thead>
            <tbody>
              {aylar.length === 0 && (
                <tr><td colSpan={5} className="muted">Veri yok</td></tr>
              )}
              {aylar.map((a) => {
                const aktif = a.ay === secili;
                return (
                  <tr key={a.ay} className={aktif ? "active" : ""} onClick={() => onSelect(a.ay)}>
                    <td>
                      {/* Klavyeyle de seçilebilsin: satır tıklaması dokunma için, düğme odak için */}
                      <button
                        type="button"
                        className="mc-pick"
                        aria-pressed={aktif}
                        title="Bu ayı seç"
                        onClick={(e) => { e.stopPropagation(); onSelect(a.ay); }}
                      >
                        {a.label}
                      </button>
                    </td>
                    <td className="num">{nf(a.adet)}</td>
                    <td className="num">{tl(a.toptanCiro)}</td>
                    <td className="num">{tl(a.perakendeCiro)}</td>
                    <td className="num"><strong>{tl(a.ciro)}</strong></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {/* Grafik kabı tablo görünümünde de bağlı kalır (hidden): useSize'ın ResizeObserver'ı
          aynı elemanı izlemeye devam eder, grafiğe dönüşte genişlik yeniden ölçülür. */}
      <div className="chart-plot" ref={ref} style={{ height: H }} hidden={table}>
        {width > 0 && (
          <svg width={width} height={H} role="img" aria-label="Son 6 ayın kişisel ciro grafiği">
            {ticks.map((t) => (
              <g key={t}>
                <line x1={padL} x2={width - padR} y1={y(t)} y2={y(t)} className="chart-grid" />
                <text x={padL - 8} y={y(t) + 4} className="chart-tick" textAnchor="end">{tickFmt(t)}</text>
              </g>
            ))}
            {hover !== null && (
              <rect x={padL + hover * band} y={padT} width={band} height={plotH} className="chart-hover-band" />
            )}
            {aylar.map((a, i) => {
              const cx = padL + i * band + band / 2;
              const x = cx - barW / 2;
              const h = a.ciro > 0 ? Math.max(2, plotH - (y(a.ciro) - padT)) : 0;
              const top = padT + plotH - h;
              const aktif = a.ay === secili;
              return (
                <g key={a.ay}>
                  {h > 0 && (
                    <path d={ustYuvarlak(x, top, barW, h, 4)} className={`mc-bar ${aktif ? "active" : ""} ${hover === i ? "hover" : ""}`} />
                  )}
                  {a.ciro > 0 && a.ciro === enYuksek && (
                    <text x={cx} y={top - 6} className="chart-label" textAnchor="middle">{tl(a.ciro)}</text>
                  )}
                  <text x={cx} y={H - padB + 16} className={`chart-tick mc-tick ${aktif ? "active" : ""}`} textAnchor="middle">{a.kisa}</text>
                  {aktif && <rect x={cx - 8} y={H - padB + 21} width={16} height={3} rx={1.5} className="mc-marker" />}
                  <rect
                    x={padL + i * band} y={padT} width={band} height={plotH + padB}
                    className="mc-hit"
                    onPointerEnter={() => setHover(i)}
                    onPointerLeave={() => setHover(null)}
                    onClick={() => onSelect(a.ay)}
                    onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onSelect(a.ay); } }}
                    tabIndex={0}
                    role="button"
                    onFocus={() => setHover(i)}
                    onBlur={() => setHover(null)}
                    aria-label={`${a.label}: ${tl(a.ciro)} ciro, ${a.adet} sipariş — seçmek için tıkla`}
                    aria-pressed={aktif}
                  />
                </g>
              );
            })}
            <line x1={padL} x2={width - padR} y1={padT + plotH} y2={padT + plotH} className="chart-axis" />
          </svg>
        )}
        {tip && hover !== null && width > 0 && (
          <div
            className="chart-tip"
            style={{ left: Math.min(width - 170, Math.max(0, padL + hover * band + band / 2 - 80)), top: 6 }}
          >
            <div className="chart-tip-head">{tip.label}</div>
            <div className="chart-tip-row"><i className="key toptan" /><b>{tl(tip.ciro)}</b><span>ciro</span></div>
            <div className="chart-tip-row"><i className="key" /><b>{nf(tip.adet)}</b><span>sipariş</span></div>
            <div className="chart-tip-row"><i className="key toptan" /><b>{tl(tip.toptanCiro)}</b><span>toptan · {tip.toptanAdet}</span></div>
            <div className="chart-tip-row"><i className="key perakende" /><b>{tl(tip.perakendeCiro)}</b><span>perakende · {tip.perakendeAdet}</span></div>
          </div>
        )}
      </div>
    </>
  );
}

/* ---------- Sayfa gövdesi ---------- */
export default function MySales({ employeeName }: { employeeName: string }) {
  const [data, setData] = useState<KisiselOzet | null>(null);
  const [ay, setAy] = useState("");          // "" = içinde bulunulan ay
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [table, setTable] = useState(false);
  const istek = useRef(0);                   // geç gelen eski yanıtlar ezmesin

  const load = useCallback(async (secim: string) => {
    const no = ++istek.current;
    setLoading(true);
    setError("");
    try {
      const r = await fetch("/api/cirom" + (secim ? `?ay=${encodeURIComponent(secim)}` : ""));
      const d = await r.json();
      if (no !== istek.current) return;
      if (d.ok) setData(d.data as KisiselOzet);
      else setError(d.error || "Veriler alınamadı.");
    } catch {
      if (no === istek.current) setError("Sunucuya ulaşılamadı.");
    } finally {
      if (no === istek.current) setLoading(false);
    }
  }, []);

  useEffect(() => { load(ay); }, [load, ay]);

  const d = data || BOS;
  const dim = loading && !!data;
  const ilkYuk = loading && !data;
  const s = d.secili, o = d.oncekiAy;
  const buAy = d.aylar.length ? d.aylar[d.aylar.length - 1] : null;
  const seciliBuAy = !buAy || buAy.ay === d.ay;
  const ayKisa = d.aylar.find((a) => a.ay === d.ay)?.kisa || "";

  // Seçili ay ↔ önceki ay kıyası
  const fark = s.ciro - o.ciro;
  const ayDelta = !data ? undefined
    : o.ciro > 0 ? `${fark >= 0 ? "+" : "−"}%${Math.round((Math.abs(fark) / o.ciro) * 100)} önceki aya göre`
    : s.ciro > 0 ? `+${tl(s.ciro)} önceki aya göre`
    : "önceki ay veri yok";
  const ayIyi = fark > 0 ? true : fark < 0 ? false : null;
  const ayOran = Math.min(1, s.ciro / Math.max(o.ciro, 1));

  const oncekiOrt = o.adet ? o.ciro / o.adet : 0;
  const ortFark = d.ortalamaSiparis - oncekiOrt;
  const ortDelta = !data ? undefined
    : oncekiOrt > 0 ? `${ortFark >= 0 ? "+" : "−"}%${Math.round((Math.abs(ortFark) / oncekiOrt) * 100)} önceki aya göre`
    : undefined;
  const ortOran = oncekiOrt > 0 ? Math.min(1, d.ortalamaSiparis / oncekiOrt) : d.ortalamaSiparis > 0 ? 1 : 0;

  const haftaOran = buAy && buAy.ciro > 0 ? Math.min(1, d.hafta.ciro / buAy.ciro) : d.hafta.ciro > 0 ? 1 : 0;
  const bugunOran = d.hafta.ciro > 0 ? Math.min(1, d.bugun.ciro / d.hafta.ciro) : d.bugun.ciro > 0 ? 1 : 0;

  const toptanPay = s.ciro > 0 ? (s.toptanCiro / s.ciro) * 100 : 0;
  const perakendePay = s.ciro > 0 ? (s.perakendeCiro / s.ciro) * 100 : 0;

  return (
    <div className="dash">
      {error && (
        <div className="notice err row">
          <span style={{ flex: 1 }}>{error}</span>
          <button type="button" className="btn small secondary" onClick={() => load(ay)}>Tekrar dene</button>
        </div>
      )}

      <div className={`kpi-row ${dim ? "loading-dim" : ""}`}>
        <StatTile
          label="Bugün"
          value={tl(d.bugun.ciro)}
          ratio={bugunOran}
          ratioLabel={`${nf(d.bugun.adet)} sipariş`}
          delta={data ? `${nf(d.bugun.toptanAdet)} toptan · ${nf(d.bugun.perakendeAdet)} perakende` : undefined}
          icon="clock"
          tone="blue"
          loading={ilkYuk}
        />
        <StatTile
          label="Son 7 Gün"
          value={tl(d.hafta.ciro)}
          ratio={haftaOran}
          ratioLabel={`${nf(d.hafta.adet)} sipariş`}
          delta={data ? `günde ort. ${tl(d.hafta.ciro / 7)}` : undefined}
          icon="activity"
          tone="amber"
          loading={ilkYuk}
        />
        <StatTile
          label={d.ayLabel}
          value={tl(s.ciro)}
          ratio={ayOran}
          ratioLabel={`${nf(s.toptanAdet)} toptan · ${nf(s.perakendeAdet)} perakende`}
          delta={ayDelta}
          deltaGood={data ? ayIyi : null}
          icon="calendar"
          tone="green"
          loading={ilkYuk}
        />
        <StatTile
          label="Ortalama Sipariş"
          value={tlOrt(d.ortalamaSiparis)}
          ratio={ortOran}
          ratioLabel={`${nf(s.adet)} sipariş${seciliBuAy ? " bu ay" : ayKisa ? ` · ${ayKisa}` : ""}`}
          delta={ortDelta}
          deltaGood={data && oncekiOrt > 0 ? (ortFark > 0 ? true : ortFark < 0 ? false : null) : null}
          icon="tag"
          tone="brand"
          loading={ilkYuk}
        />
      </div>

      <div className={`dash-grid ${dim ? "loading-dim" : ""}`}>
        {/* Aylık ciro */}
        <section className="card span-7">
          <div className="card-head mc-head">
            <span className="card-head-icon"><Icon name="bar-chart" size={18} /></span>
            <div>
              <h2>Aylık Ciro</h2>
              <span className="card-head-sub">Son 6 ay · seçili: {d.ayLabel}</span>
            </div>
            <span className="spacer" />
            <div className="card-head-actions">
              <button type="button" className="btn xs ghost" onClick={() => setTable((t) => !t)} aria-pressed={table}>
                {table ? "Grafik" : "Tablo"}
              </button>
              <button type="button" className="btn icon ghost small" onClick={() => load(ay)} title="Yenile" aria-label="Yenile">
                <Icon name="refresh" size={16} className={loading ? "spin" : undefined} />
              </button>
            </div>
            {d.aylar.length > 0 && (
              <div className="seg mc-seg" role="group" aria-label="Ay seç">
                {d.aylar.map((a) => (
                  <button
                    key={a.ay}
                    type="button"
                    aria-pressed={a.ay === d.ay}
                    className={a.ay === d.ay ? "active" : ""}
                    onClick={() => setAy(a.ay)}
                    title={a.label}
                  >
                    {a.kisa}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className={ilkYuk ? "loading-dim" : ""}>
            <MonthChart aylar={d.aylar} secili={d.ay} table={table} onSelect={setAy} />
          </div>
        </section>

        {/* Son 14 gün */}
        <section className="card span-5">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="activity" size={18} /></span>
            <div>
              <h2>Son 14 Gün</h2>
              <span className="card-head-sub">Günlük ciro · toptan / perakende</span>
            </div>
          </div>
          <OrderFlowChart seri={d.seri14} mode="ciro" loading={ilkYuk} />
        </section>

        {/* Ay özeti */}
        <section className="card span-5">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="pie" size={18} /></span>
            <div>
              <h2>{seciliBuAy ? "Bu Ayın Özeti" : `${d.ayLabel} Özeti`}</h2>
              <span className="card-head-sub">{d.ayLabel} · {nf(s.adet)} sipariş</span>
            </div>
          </div>
          <ul className="ms-rows">
            <li>
              <span className="lbl"><i className="swatch toptan" /> Toptan</span>
              <span className="val">{nf(s.toptanAdet)} sipariş · {tl(s.toptanCiro)}</span>
            </li>
            <li>
              <span className="lbl"><i className="swatch perakende" /> Perakende</span>
              <span className="val">{nf(s.perakendeAdet)} sipariş · {tl(s.perakendeCiro)}</span>
            </li>
            <li>
              <span className="lbl"><Icon name="star" size={14} /> En iyi gün</span>
              <span className="val">{d.enIyiGun ? <>{d.enIyiGun.label} · {tl(d.enIyiGun.ciro)}</> : "—"}</span>
            </li>
            <li>
              <span className="lbl"><Icon name="calendar" size={14} /> Önceki ay</span>
              <span className="val">{tl(o.ciro)} <small>· {nf(o.adet)} sipariş</small></span>
            </li>
          </ul>
          <div className="ms-split">
            <div className="chart-legend">
              <span className="chart-key"><i className="swatch toptan" /> Toptan <b>%{Math.round(toptanPay)}</b></span>
              <span className="chart-key"><i className="swatch perakende" /> Perakende <b>%{Math.round(perakendePay)}</b></span>
            </div>
            <div className="rep-bar" role="img" aria-label={`Toptan %${Math.round(toptanPay)}, perakende %${Math.round(perakendePay)}`}>
              <span className="seg toptan" style={{ width: `${toptanPay}%` }} />
              <span className="seg perakende" style={{ width: `${perakendePay}%` }} />
            </div>
          </div>
        </section>

        {/* Siparişlerim */}
        <section className="card span-7 pad-0">
          <div className="card-head" style={{ margin: 0 }}>
            <span className="card-head-icon"><Icon name="list" size={18} /></span>
            <div>
              <h2>Siparişlerim — {d.ayLabel}</h2>
              <span className="card-head-sub">
                {!d.blob && data ? "Depo bağlı değil" : d.sonSiparisler.length >= 30 ? "Son 30 kayıt" : `${nf(d.sonSiparisler.length)} sipariş · en yeni üstte`}
              </span>
            </div>
            <span className="spacer" />
            <Link href="/panel/siparisler" className="btn small secondary">Tüm siparişler</Link>
          </div>
          {ilkYuk ? (
            <ul className="recent recent-skel" aria-busy="true" aria-label="Siparişler yükleniyor">
              {[0, 1, 2, 3].map((i) => (
                <li key={i}>
                  <span className="recent-row">
                    <span className="recent-avatar" />
                    <span className="recent-main">
                      <span className="skeleton" style={{ width: "55%", height: 13, marginBottom: 6 }} />
                      <span className="skeleton" style={{ width: "38%", height: 11 }} />
                    </span>
                  </span>
                </li>
              ))}
            </ul>
          ) : data && d.sonSiparisler.length === 0 ? (
            <div className="empty">
              <div className="empty-icon"><Icon name="inbox" size={22} /></div>
              <strong>{d.ayLabel} içinde senin adına sipariş yok</strong>
              {d.blob ? "Girdiğin toptan ve perakende siparişler burada listelenir." : "Kalıcı depolama bağlandığında siparişler burada görünür."}
            </div>
          ) : (
            <ul className="recent">
              {d.sonSiparisler.map((sp) => (
                <li key={`${sp.tur}-${sp.orderId}`}>
                  <Link href={sp.href} className="recent-row">
                    <span className={`recent-avatar ${sp.tur === "perakende" ? "retail" : ""}`}>
                      {(sp.musteri || "?").slice(0, 1).toLocaleUpperCase("tr-TR")}
                    </span>
                    <span className="recent-main">
                      <span className="recent-title">{sp.musteri || "—"}</span>
                      <span className="recent-sub">{sp.orderId} · {sp.tur === "toptan" ? "Toptan" : "Perakende"} · {gunKisa(sp.dateKey)} {saat(sp.createdAt)}</span>
                    </span>
                    <span className="recent-amt num">{tl(sp.tutar)}</span>
                    <span className={`badge ${sp.statusKind}`}>{sp.status}</span>
                  </Link>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>

      <p className="ms-foot">
        <Icon name="shield" size={14} />
        <span>
          Bu ekranda yalnızca senin adına girilen siparişler sayılır; başka çalışanların rakamları görünmez.
          {" "}<span className="text-2">Hesap: {employeeName}</span>
        </span>
      </p>
    </div>
  );
}
