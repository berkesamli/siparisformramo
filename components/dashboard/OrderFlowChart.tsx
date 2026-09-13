"use client";

// 14 günlük sipariş akışı — gruplu sütun (toptan / perakende), fare ile
// gün vurgusu + tek tooltip (her iki seri), tablo görünümü ikizi.

import { useMemo, useState } from "react";
import type { SeriGun } from "@/lib/dashboard";
import { useSize } from "./useSize";

const fmtTL = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });

function niceMax(v: number): number {
  if (v <= 0) return 4;
  const p = Math.pow(10, Math.floor(Math.log10(v)));
  const m = v / p;
  const step = m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10;
  const top = step * p;
  return top === v ? v + (top / 4 || 1) : top;
}

export default function OrderFlowChart({
  seri,
  mode,
  loading,
}: {
  seri: SeriGun[];
  mode: "adet" | "ciro";
  loading?: boolean;
}) {
  const { ref, width } = useSize<HTMLDivElement>();
  const [hover, setHover] = useState<number | null>(null);
  const [table, setTable] = useState(false);

  const H = 230;
  const padL = 40, padR = 10, padT = 18, padB = 34;
  const plotW = Math.max(0, width - padL - padR);
  const plotH = H - padT - padB;
  const n = seri.length || 1;
  const band = plotW / n;

  const vals = useMemo(() => seri.map((g) => (mode === "adet" ? [g.toptan, g.perakende] : [g.toptanCiro, g.perakendeCiro])), [seri, mode]);
  const maxV = niceMax(Math.max(0, ...vals.flat()));
  const ticks = [0, 0.25, 0.5, 0.75, 1].map((t) => t * maxV);
  const y = (v: number) => padT + plotH - (v / maxV) * plotH;
  const barW = Math.min(24, Math.max(6, band * 0.3));
  const gap = 2;
  const every = Math.max(1, Math.ceil(40 / Math.max(1, band)));
  const showLabel = (i: number) => (n - 1 - i) % every === 0;
  const fmt = (v: number) => (mode === "adet" ? String(Math.round(v)) : fmtTL(v));
  const tickFmt = (v: number) => (mode === "adet" ? String(Math.round(v)) : v >= 1000 ? `${Math.round(v / 1000)}k` : String(Math.round(v)));

  const maxT = Math.max(...vals.map((v) => v[0]));
  const maxP = Math.max(...vals.map((v) => v[1]));

  const toplamT = vals.reduce((s, v) => s + v[0], 0);
  const toplamP = vals.reduce((s, v) => s + v[1], 0);

  return (
    <div className={`chart ${loading ? "loading-dim" : ""}`}>
      <div className="chart-legend">
        <span className="chart-key"><i className="swatch toptan" /> Toptan <b>{fmt(toplamT)}</b></span>
        <span className="chart-key"><i className="swatch perakende" /> Perakende <b>{fmt(toplamP)}</b></span>
        <span className="spacer" />
        <button type="button" className="btn xs ghost" onClick={() => setTable((t) => !t)} aria-pressed={table}>
          {table ? "Grafik" : "Tablo"}
        </button>
      </div>

      {table ? (
        <div className="table-wrap">
          <table className="chart-table">
            <thead>
              <tr><th>Gün</th><th className="num">Toptan</th><th className="num">Perakende</th><th className="num">Toplam</th></tr>
            </thead>
            <tbody>
              {seri.map((g, i) => (
                <tr key={g.date}>
                  <td>{g.label} <span className="muted">{g.gun}</span></td>
                  <td className="num">{fmt(vals[i][0])}</td>
                  <td className="num">{fmt(vals[i][1])}</td>
                  <td className="num"><strong>{fmt(vals[i][0] + vals[i][1])}</strong></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="chart-plot" ref={ref} style={{ height: H }}>
          {width > 0 && (
            <svg width={width} height={H} role="img" aria-label="Son 14 günün sipariş grafiği">
              {ticks.map((t) => (
                <g key={t}>
                  <line x1={padL} x2={width - padR} y1={y(t)} y2={y(t)} className="chart-grid" />
                  <text x={padL - 8} y={y(t) + 4} className="chart-tick" textAnchor="end">{tickFmt(t)}</text>
                </g>
              ))}
              {hover !== null && (
                <rect x={padL + hover * band} y={padT} width={band} height={plotH} className="chart-hover-band" />
              )}
              {seri.map((g, i) => {
                const [t, p] = vals[i];
                const cx = padL + i * band + band / 2;
                const xT = cx - barW - gap / 2;
                const xP = cx + gap / 2;
                const hT = Math.max(t > 0 ? 2 : 0, plotH - (y(t) - padT));
                const hP = Math.max(p > 0 ? 2 : 0, plotH - (y(p) - padT));
                return (
                  <g key={g.date}>
                    {t > 0 && <rect x={xT} y={y(t)} width={barW} height={hT} rx={3} className="chart-bar toptan" />}
                    {p > 0 && <rect x={xP} y={y(p)} width={barW} height={hP} rx={3} className="chart-bar perakende" />}
                    {t > 0 && t === maxT && (
                      <text x={xT + barW / 2} y={y(t) - 5} className="chart-label" textAnchor="middle">{fmt(t)}</text>
                    )}
                    {p > 0 && p === maxP && (
                      <text x={xP + barW / 2} y={y(p) - 5} className="chart-label" textAnchor="middle">{fmt(p)}</text>
                    )}
                    {showLabel(i) && (
                      <text x={cx} y={H - padB + 16} className="chart-tick" textAnchor="middle">{g.label}</text>
                    )}
                    {showLabel(i) && band >= 44 && (
                      <text x={cx} y={H - padB + 29} className="chart-tick faint" textAnchor="middle">{g.gun}</text>
                    )}
                    <rect
                      x={padL + i * band} y={padT} width={band} height={plotH + padB}
                      fill="transparent"
                      onPointerEnter={() => setHover(i)}
                      onPointerLeave={() => setHover(null)}
                      tabIndex={0}
                      onFocus={() => setHover(i)}
                      onBlur={() => setHover(null)}
                      aria-label={`${g.label}: toptan ${fmt(t)}, perakende ${fmt(p)}`}
                    />
                  </g>
                );
              })}
              <line x1={padL} x2={width - padR} y1={padT + plotH} y2={padT + plotH} className="chart-axis" />
            </svg>
          )}
          {hover !== null && seri[hover] && width > 0 && (
            <div
              className="chart-tip"
              style={{
                left: Math.min(width - 170, Math.max(0, padL + hover * band + band / 2 - 80)),
                top: 6,
              }}
            >
              <div className="chart-tip-head">{seri[hover].label} · {seri[hover].gun}</div>
              <div className="chart-tip-row"><i className="key toptan" /><b>{fmt(vals[hover][0])}</b><span>Toptan</span></div>
              <div className="chart-tip-row"><i className="key perakende" /><b>{fmt(vals[hover][1])}</b><span>Perakende</span></div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
