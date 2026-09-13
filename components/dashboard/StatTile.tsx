"use client";

import Link from "next/link";
import Icon, { type IconName } from "@/components/shell/Icon";

/**
 * KPI kutusu: halka ölçer + değer + etiket + karşılaştırma satırı.
 * ratio 0–1 halkanın doluluğu; tone halkanın rengi.
 */
export default function StatTile({
  label,
  value,
  ratio,
  ratioLabel,
  delta,
  deltaGood,
  icon,
  href,
  tone = "brand",
  loading,
}: {
  label: string;
  value: string;
  ratio: number;
  ratioLabel?: string;
  delta?: string;
  deltaGood?: boolean | null;
  icon: IconName;
  href?: string;
  tone?: "brand" | "blue" | "green" | "amber" | "red";
  loading?: boolean;
}) {
  const r = 26;
  const c = 2 * Math.PI * r;
  const p = Math.max(0, Math.min(1, Number.isFinite(ratio) ? ratio : 0));
  const body = (
    <>
      <div className={`stat-ring tone-${tone}`} aria-hidden>
        <svg width="68" height="68" viewBox="0 0 68 68" style={{ width: "100%", height: "100%" }}>
          <circle cx="34" cy="34" r={r} className="stat-ring-track" />
          {p > 0 && (
            <circle
              cx="34" cy="34" r={r}
              className="stat-ring-fill"
              strokeDasharray={`${c * p} ${c}`}
              transform="rotate(-90 34 34)"
            />
          )}
        </svg>
        <span className="stat-ring-icon"><Icon name={icon} size={18} /></span>
      </div>
      <div className="stat-body">
        <span className="stat-label">{label}</span>
        <span className={`stat-value ${loading ? "skeleton" : ""}`}>{loading ? " " : value}</span>
        <span className="stat-foot">
          {delta && (
            <span className={`stat-delta ${deltaGood === true ? "good" : deltaGood === false ? "bad" : ""}`}>
              {delta}
            </span>
          )}
          {ratioLabel && <span className="stat-ratio">{ratioLabel}</span>}
        </span>
      </div>
    </>
  );
  return href ? (
    <Link href={href} className="stat card" title={label}>{body}</Link>
  ) : (
    <div className="stat card">{body}</div>
  );
}
