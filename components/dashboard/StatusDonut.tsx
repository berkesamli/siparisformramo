"use client";

// Durum dağılımı halkası — son 14 gün. Dilimler arasında 2 px yüzey boşluğu,
// ortada toplam; lejant aynı zamanda tablo ikizidir (adet + yüzde).

import { useState } from "react";
import { STATUS_LABELS, type OrderStatus } from "@/lib/orders";
import type { RetailStatus } from "@/data/perakende";

interface Dilim { key: string; label: string; value: number; cls: string; }

const TOPTAN_SIRA: OrderStatus[] = ["olusturuldu", "hazirlaniyor", "yarim", "tamamlandi", "iptal"];
const TOPTAN_CLS: Record<OrderStatus, string> = { olusturuldu: "st-olusturuldu", hazirlaniyor: "st-hazirlaniyor", yarim: "st-yarim", tamamlandi: "st-tamamlandi", iptal: "st-iptal" };
const PERAKENDE_SIRA: RetailStatus[] = ["Beklemede", "Hazırlanıyor", "Hazır", "Teslim Edildi", "İptal"];
const PERAKENDE_CLS: Record<RetailStatus, string> = { Beklemede: "st-olusturuldu", "Hazırlanıyor": "st-hazirlaniyor", "Hazır": "st-tamamlandi", "Teslim Edildi": "st-teslim", "İptal": "st-iptal" };

export default function StatusDonut({
  toptan,
  perakende,
  loading,
}: {
  toptan: Record<OrderStatus, number>;
  perakende: Record<RetailStatus, number>;
  loading?: boolean;
}) {
  const [kanal, setKanal] = useState<"toptan" | "perakende">("toptan");
  const dilimler: Dilim[] =
    kanal === "toptan"
      ? TOPTAN_SIRA.map((k) => ({ key: k, label: STATUS_LABELS[k], value: toptan[k] || 0, cls: TOPTAN_CLS[k] }))
      : PERAKENDE_SIRA.map((k) => ({ key: k, label: k, value: perakende[k] || 0, cls: PERAKENDE_CLS[k] }));
  const toplam = dilimler.reduce((s, d) => s + d.value, 0);

  const R = 54, SW = 14;
  const C = 2 * Math.PI * R;
  const gapPx = toplam > 0 && dilimler.filter((d) => d.value > 0).length > 1 ? 2 : 0;
  let offset = 0;
  const arcs = dilimler.map((d) => {
    const len = toplam > 0 ? (d.value / toplam) * C : 0;
    const a = { ...d, len: Math.max(0, len - gapPx), start: offset };
    offset += len;
    return a;
  });

  return (
    <div className={`donut ${loading ? "loading-dim" : ""}`}>
      <div className="seg" style={{ alignSelf: "flex-start" }}>
        <button type="button" className={kanal === "toptan" ? "active" : ""} onClick={() => setKanal("toptan")}>Toptan</button>
        <button type="button" className={kanal === "perakende" ? "active" : ""} onClick={() => setKanal("perakende")}>Perakende</button>
      </div>
      <div className="donut-body">
        <div className="donut-svg">
          <svg width="150" height="150" viewBox="0 0 150 150" role="img" aria-label="Sipariş durumu dağılımı">
            <circle cx="75" cy="75" r={R} className="donut-track" strokeWidth={SW} />
            {toplam > 0 && arcs.filter((a) => a.len > 0).map((a) => (
              <circle
                key={a.key}
                cx="75" cy="75" r={R}
                className={`donut-arc ${a.cls}`}
                strokeWidth={SW}
                strokeDasharray={`${a.len} ${C}`}
                strokeDashoffset={-a.start}
                transform="rotate(-90 75 75)"
              >
                <title>{a.label}: {a.value}</title>
              </circle>
            ))}
          </svg>
          <div className="donut-center">
            <strong>{toplam}</strong>
            <span>14 günde</span>
          </div>
        </div>
        <ul className="donut-legend" aria-label="Durum dağılımı tablosu">
          {dilimler.map((d) => (
            <li key={d.key}>
              <i className={`swatch ${d.cls}`} />
              <span className="lbl">{d.label}</span>
              <b className="num">{d.value}</b>
              <span className="pct num">{toplam ? Math.round((d.value / toplam) * 100) : 0}%</span>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}
