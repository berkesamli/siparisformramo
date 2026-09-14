"use client";

// Gösterge panelindeki kompakt "Satışlarım" kartı: bugün / son 7 gün / bu ay
// ve 14 günlük mini çizgi. Yalnızca oturumdaki çalışanın kendi siparişleri.

import { useEffect, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import type { KisiselOzet } from "@/lib/personal-sales";
import { useSize } from "@/components/dashboard/useSize";

const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });
const nf = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");

function Sparkline({ seri }: { seri: KisiselOzet["seri14"] }) {
  const { ref, width } = useSize<HTMLDivElement>();
  const [hover, setHover] = useState<number | null>(null);

  const H = 64;
  const padT = 6, padB = 4, padX = 4;
  const vals = seri.map((g) => (Number(g.toptanCiro) || 0) + (Number(g.perakendeCiro) || 0));
  const n = vals.length;
  const max = Math.max(1, ...vals);
  const plotW = Math.max(0, width - padX * 2);
  const plotH = H - padT - padB;
  const px = (i: number) => padX + (n > 1 ? (i / (n - 1)) * plotW : plotW / 2);
  const py = (v: number) => padT + plotH - (v / max) * plotH;
  const pts = vals.map((v, i) => [px(i), py(v)] as const);
  const line = pts.map(([x, y], i) => `${i ? "L" : "M"}${x.toFixed(1)},${y.toFixed(1)}`).join(" ");
  const area = n ? `${line} L${pts[n - 1][0].toFixed(1)},${padT + plotH} L${pts[0][0].toFixed(1)},${padT + plotH} Z` : "";
  const son = n ? pts[n - 1] : null;
  const tip = hover !== null ? seri[hover] : null;

  const onMove = (e: React.PointerEvent<SVGSVGElement>) => {
    if (!n || plotW <= 0) return;
    const r = e.currentTarget.getBoundingClientRect();
    const x = e.clientX - r.left - padX;
    const i = Math.round((x / plotW) * (n - 1));
    setHover(Math.max(0, Math.min(n - 1, i)));
  };

  return (
    <div className="ms-spark" ref={ref}>
      {width > 0 && n > 0 && (
        <svg width={width} height={H} viewBox={`0 0 ${width} ${H}`} role="img" aria-label="Son 14 günün günlük ciro çizgisi"
          onPointerMove={onMove} onPointerLeave={() => setHover(null)}>
          <line x1={padX} x2={width - padX} y1={padT + plotH} y2={padT + plotH} className="ms-spark-base" />
          <path d={area} className="ms-spark-area" />
          <path d={line} className="ms-spark-line" />
          {hover !== null && (
            <>
              <line x1={pts[hover][0]} x2={pts[hover][0]} y1={padT} y2={padT + plotH} className="ms-spark-cursor" />
              <circle cx={pts[hover][0]} cy={pts[hover][1]} r={3.5} className="ms-spark-dot" />
            </>
          )}
          {son && hover === null && <circle cx={son[0]} cy={son[1]} r={4} className="ms-spark-dot" />}
        </svg>
      )}
      {tip && hover !== null && width > 0 && (
        <div className="chart-tip" style={{ left: Math.min(width - 170, Math.max(0, pts[hover][0] - 80)), top: -8 }}>
          <div className="chart-tip-head">{tip.label} · {tip.gun}</div>
          <div className="chart-tip-row"><i className="key toptan" /><b>{tl(vals[hover])}</b><span>ciro · {tip.toptan + tip.perakende} sipariş</span></div>
        </div>
      )}
      <div className="ms-spark-foot">
        <span>{seri[0]?.label || ""}</span>
        <span>Son 14 gün · günlük ciro</span>
        <span>{seri[n - 1]?.label || ""}</span>
      </div>
    </div>
  );
}

export default function MySalesCard() {
  const [data, setData] = useState<KisiselOzet | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let iptal = false;
    (async () => {
      try {
        const r = await fetch("/api/cirom");
        const d = await r.json();
        if (iptal) return;
        if (d.ok) setData(d.data as KisiselOzet);
        else setError(d.error || "Satış özeti alınamadı.");
      } catch {
        if (!iptal) setError("Satış özeti alınamadı.");
      } finally {
        if (!iptal) setLoading(false);
      }
    })();
    return () => { iptal = true; };
  }, []);

  // ?ay verilmeyince sunucu içinde bulunulan ayı "secili" olarak döndürür.
  const buAy = data ? data.secili : null;

  return (
    <section className="card span-5">
      <div className="card-head">
        <span className="card-head-icon"><Icon name="trending-up" size={18} /></span>
        <div>
          <h2>Satışlarım</h2>
          <span className="card-head-sub">Sadece senin siparişlerin</span>
        </div>
        <span className="spacer" />
        <div className="card-head-actions">
          <Link href="/panel/satislarim" className="btn small secondary">
            Ayrıntılar <Icon name="arrow-up-right" size={14} />
          </Link>
        </div>
      </div>

      {loading ? (
        <>
          <div className="ms-stats" aria-busy>
            {[0, 1, 2].map((i) => (
              <div className="ms-stat" key={i}>
                <span className="skeleton" style={{ width: "50%", height: 11 }} />
                <span className="ms-stat-value skeleton" />
                <span className="skeleton" style={{ width: "40%", height: 11 }} />
              </div>
            ))}
          </div>
          <div className="ms-spark"><div className="skeleton ms-spark-skel" /></div>
        </>
      ) : error || !data ? (
        <p className="ms-err">{error || "Satış özeti alınamadı."}</p>
      ) : (
        <>
          <div className="ms-stats">
            <div className="ms-stat">
              <span className="ms-stat-label">Bugün</span>
              <span className="ms-stat-value">{tl(data.bugun.ciro)}</span>
              <span className="ms-stat-sub">{nf(data.bugun.adet)} sipariş</span>
            </div>
            <div className="ms-stat">
              <span className="ms-stat-label">Son 7 gün</span>
              <span className="ms-stat-value">{tl(data.hafta.ciro)}</span>
              <span className="ms-stat-sub">{nf(data.hafta.adet)} sipariş</span>
            </div>
            <div className="ms-stat">
              <span className="ms-stat-label">Bu ay</span>
              <span className="ms-stat-value">{tl(buAy ? buAy.ciro : 0)}</span>
              <span className="ms-stat-sub">{nf(buAy ? buAy.adet : 0)} sipariş</span>
            </div>
          </div>
          <Sparkline seri={data.seri14} />
        </>
      )}
    </section>
  );
}
