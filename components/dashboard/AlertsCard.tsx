"use client";

import Link from "next/link";
import Icon, { type IconName } from "@/components/shell/Icon";
import type { Uyari } from "@/lib/dashboard";

const ICON: Record<Uyari["tip"], IconName> = { kontrol: "eye", stok: "package", kur: "dollar", cek: "file-text", blob: "info", acik: "activity" };

export default function AlertsCard({ uyarilar, loading }: { uyarilar: Uyari[]; loading?: boolean }) {
  if (!loading && uyarilar.length === 0) {
    return (
      <div className="empty" style={{ padding: "22px 12px" }}>
        <div className="empty-icon" style={{ color: "var(--success)", background: "var(--success-soft)" }}><Icon name="check-circle" size={22} /></div>
        <strong>Her şey yolunda</strong>
        Bekleyen kontrol, eski stok verisi veya vadesi geçmiş evrak yok.
      </div>
    );
  }
  return (
    <ul className={`alerts ${loading ? "loading-dim" : ""}`}>
      {uyarilar.map((u, i) => (
        <li key={i} className={`alert-row ${u.seviye}`}>
          <Link href={u.href} className="alert-link">
            <span className="alert-icon"><Icon name={ICON[u.tip]} size={17} /></span>
            <span className="alert-main">
              <strong>{u.baslik}</strong>
              <span>{u.metin}</span>
            </span>
            <Icon name="chevron-right" size={16} className="alert-chev" />
          </Link>
        </li>
      ))}
    </ul>
  );
}
