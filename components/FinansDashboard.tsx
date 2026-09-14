"use client";

// Finans genel bakış — kasa özeti, aylık tahsilat/gider/kâr çubukları,
// vadesi yaklaşan çekler. Yalnızca aylık özet dosyalarını okur (hızlı).

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import type { FinansOzet, SubeKey } from "@/lib/finans-ozet";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
const fmt0 = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });

const AY_ADI = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];

/** Çubuk yüksekliği (px) — grafik alanı bu ölçüye göre ölçeklenir. */
const BAR_H = 130;

interface VadeSatir {
  id: string;
  tur: string;
  kind: string;
  vade: string;
  tutar: number;
  kimden: string;
  branch: string;
  gecmis: boolean;
}

interface PortfoyOzet {
  alinanAdet: number;
  alinanToplam: number;
  verilenAdet: number;
  verilenToplam: number;
}

/**
 * Bir kutunun genişliğini izler (ResizeObserver). Grafik dar ekranda
 * (telefon / daraltılmış kenar çubuğu) yatay kaydırma yerine sıkışık moda geçer.
 */
function useWidth(): [(el: HTMLDivElement | null) => void, number] {
  const [w, setW] = useState(0);
  const obs = useRef<ResizeObserver | null>(null);
  const ref = useCallback((el: HTMLDivElement | null) => {
    obs.current?.disconnect();
    obs.current = null;
    if (!el || typeof ResizeObserver === "undefined") return;
    setW(Math.round(el.clientWidth));
    const ro = new ResizeObserver((entries) => {
      const cw = entries[0]?.contentRect.width;
      if (cw) setW(Math.round(cw));
    });
    ro.observe(el);
    obs.current = ro;
  }, []);
  return [ref, w];
}

export default function FinansDashboard() {
  const [months, setMonths] = useState<string[]>([]);
  const [ozetler, setOzetler] = useState<FinansOzet[]>([]);
  const [vadeler, setVadeler] = useState<VadeSatir[]>([]);
  const [portfoy, setPortfoy] = useState<PortfoyOzet | null>(null);
  const [sube, setSube] = useState<"" | "ankara" | "istanbul">("");
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [chartRef, chartW] = useWidth();

  useEffect(() => {
    fetch("/api/finans/ozet")
      .then((r) => r.json())
      .then((d) => {
        if (d.ok) {
          setMonths(d.months || []);
          setOzetler(d.ozetler || []);
          setVadeler(d.vadesiYaklasan || []);
          setPortfoy(d.portfoyOzet || null);
        } else setErr(d.error || "Yüklenemedi");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, []);

  // Ay → {tahsilat, gider, kar} — şube filtresine göre
  const seriler = useMemo(() => {
    const map = new Map(ozetler.map((o) => [o.month, o]));
    return months.map((m) => {
      const o = map.get(m);
      let tahsilat = 0, gider = 0;
      if (o) {
        const keys: SubeKey[] = sube ? [sube] : ["ankara", "istanbul", "belirsiz"];
        for (const k of keys) {
          const s = o.sube[k];
          if (!s) continue;
          tahsilat += s.tahsilatToplam;
          gider += s.giderToplam;
        }
      }
      return { month: m, tahsilat, gider, kar: tahsilat - gider };
    });
  }, [months, ozetler, sube]);

  const buAy = seriler[seriler.length - 1];
  const maxDeger = Math.max(1, ...seriler.map((s) => Math.max(s.tahsilat, s.gider)));
  // Dar kutu (telefon): sütun aralığı ve yazılar küçülür, ay/yıl iki satıra iner.
  const compact = chartW > 0 && chartW < 520;

  if (loading) return <p className="muted">Yükleniyor…</p>;
  if (err) return <div className="notice err">{err}</div>;

  return (
    <div>
      <div className="seg no-print" role="tablist" aria-label="Şube" style={{ marginBottom: 14 }}>
        {(["", "ankara", "istanbul"] as const).map((s) => (
          <button
            key={s || "tum"}
            type="button"
            role="tab"
            aria-selected={sube === s}
            className={sube === s ? "active" : ""}
            onClick={() => setSube(s)}
          >
            {s === "" ? "Tümü" : s === "ankara" ? "Ankara" : "İstanbul"}
          </button>
        ))}
      </div>

      {/* Bu ay kartları */}
      <div className="cari-cards">
        <div className="cari-card">
          <span>Bu Ay Tahsilat</span>
          <strong style={{ color: "var(--success)" }}>₺{fmt(buAy?.tahsilat || 0)}</strong>
        </div>
        <div className="cari-card">
          <span>Bu Ay Gider</span>
          <strong style={{ color: "var(--error)" }}>₺{fmt(buAy?.gider || 0)}</strong>
        </div>
        <div className={`cari-card ${(buAy?.kar || 0) < 0 ? "borc" : ""}`}>
          <span>Bu Ay Kasa Kârı</span>
          <strong style={{ color: (buAy?.kar || 0) >= 0 ? "var(--success)" : "var(--error)" }}>
            ₺{fmt(buAy?.kar || 0)}
          </strong>
        </div>
        {portfoy && (
          <div className="cari-card">
            <span>Çek/Senet Portföyü</span>
            <strong>₺{fmt(portfoy.alinanToplam)}</strong>
            <span style={{ fontSize: 12 }}>{portfoy.alinanAdet} alınan kayıt</span>
          </div>
        )}
      </div>

      {/* Aylık çubuklar */}
      <div className="card">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="bar-chart" size={16} /></span>
          <div>
            <h2>Son 12 Ay — Tahsilat / Gider</h2>
          </div>
        </div>
        <div
          ref={chartRef}
          style={{ display: "flex", gap: compact ? 4 : 10, alignItems: "flex-end", padding: "8px 0 0", minWidth: 0 }}
        >
          {seriler.map((s) => {
            const [yy, mm] = s.month.split("-");
            const ay = AY_ADI[Number(mm) - 1];
            const karMetin = s.tahsilat || s.gider ? `${s.kar >= 0 ? "+" : ""}${fmt0(s.kar / 1000)}k` : "";
            return (
              <div
                key={s.month}
                title={`${ay} ${yy} — Tahsilat ₺${fmt(s.tahsilat)} · Gider ₺${fmt(s.gider)} · Kâr ₺${fmt(s.kar)}`}
                style={{ flex: "1 1 0", minWidth: 0, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}
              >
                <div style={{ display: "flex", gap: compact ? 2 : 3, alignItems: "flex-end", justifyContent: "center", height: BAR_H, width: "100%", maxWidth: 40 }}>
                  <div
                    title={`Tahsilat ₺${fmt(s.tahsilat)}`}
                    style={{ flex: "1 1 0", maxWidth: 16, minWidth: 3, height: Math.max(2, (s.tahsilat / maxDeger) * BAR_H),
                      background: "var(--success)", borderRadius: "3px 3px 0 0",
                      WebkitPrintColorAdjust: "exact", printColorAdjust: "exact" }}
                  />
                  <div
                    title={`Gider ₺${fmt(s.gider)}`}
                    style={{ flex: "1 1 0", maxWidth: 16, minWidth: 3, height: Math.max(2, (s.gider / maxDeger) * BAR_H),
                      background: "var(--error)", borderRadius: "3px 3px 0 0", opacity: 0.85,
                      WebkitPrintColorAdjust: "exact", printColorAdjust: "exact" }}
                  />
                </div>
                <span className="muted" style={{ fontSize: compact ? 10.5 : 11.5, lineHeight: 1.15, textAlign: "center", whiteSpace: "nowrap" }}>
                  {compact ? (<>{ay}<br />{yy.slice(2)}</>) : `${ay} ${yy.slice(2)}`}
                </span>
                <span className="num" style={{ fontSize: compact ? 10 : 11, fontWeight: 600, whiteSpace: "nowrap",
                  color: s.kar >= 0 ? "var(--success)" : "var(--error)" }}>
                  {karMetin}
                </span>
              </div>
            );
          })}
        </div>
        <p className="muted" style={{ margin: "10px 0 0", fontSize: 12.5 }}>
          Yeşil: tahsilat · Kırmızı: gider · Alt satır: kasa kârı (bin ₺). Kâr,
          kasa bazlıdır (tahsilat − gider); çek portföyü tahsil edildikçe eklenir.
        </p>
      </div>

      {/* Vadesi yaklaşan çekler */}
      <div className="card">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="clock" size={16} /></span>
          <div>
            <h2>Vadesi Yaklaşan Çek/Senet (30 gün)</h2>
          </div>
          <div className="card-head-actions">
            <Link href="/panel/finans/ceksenet" className="btn small secondary no-print">
              Portföye Git →
            </Link>
          </div>
        </div>
        {vadeler.length === 0 ? (
          <div className="empty" style={{ padding: "18px 12px" }}>
            <div className="empty-icon"><Icon name="check-circle" size={22} /></div>
            Önümüzdeki 30 günde vadesi dolan kayıt yok.
          </div>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Vade</th>
                  <th>Tür</th>
                  <th>Kimden / Kime</th>
                  <th>Şube</th>
                  <th className="num">Tutar</th>
                </tr>
              </thead>
              <tbody>
                {vadeler.map((v) => (
                  <tr key={v.id}>
                    <td style={{ fontWeight: 600, whiteSpace: "nowrap", color: v.gecmis ? "var(--error)" : undefined }}>
                      {v.vade.split("-").reverse().join(".")} {v.gecmis && "⚠"}
                    </td>
                    <td style={{ fontSize: 12.5 }}>
                      {v.tur === "alinan" ? "Alınan" : "Verilen"} {v.kind === "cek" ? "çek" : "senet"}
                    </td>
                    <td>{v.kimden}</td>
                    <td style={{ fontSize: 12.5 }}>{v.branch === "istanbul" ? "İST" : "ANK"}</td>
                    <td className="num" style={{ fontWeight: 700 }}>₺{fmt(v.tutar)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
