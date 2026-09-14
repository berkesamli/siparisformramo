"use client";

// Müşteri / bayi ana sayfası: hızlı erişim kutuları + katalog/stok özeti.

import { useEffect, useState, type CSSProperties } from "react";
import Icon from "@/components/shell/Icon";
import type { CustomerDashboard as CD } from "@/lib/dashboard";
import StatTile from "./StatTile";
import QuickActions, { type QuickTile } from "./QuickActions";

const nf = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");

// "Sipariş Hattı" kartı: telefon numarası bir metrik değil — StatTile'ın 28–32px değer stili
// 14 karakteri sığdıramaz. Aynı .stat/.card sınıfları kullanılır; değer boyutu kart genişliğine
// göre (container query birimi) 14–20px arasında ölçeklenir, sadece "0850 305 | 75 45" arasında kırılabilir.
const PHONE_BODY: CSSProperties = { flex: "1 1 auto", width: "100%", minWidth: 0, containerType: "inline-size" };
const PHONE_VALUE: CSSProperties = { fontSize: "clamp(14px, 12cqw, 20px)", lineHeight: 1.15 };
const PHONE_LINE: CSSProperties = { whiteSpace: "normal" };

export default function CustomerDashboard({ tiles }: { tiles: QuickTile[] }) {
  const [data, setData] = useState<CD | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch("/api/dashboard")
      .then((r) => r.json())
      .then((d) => { if (d.ok) setData(d.data as CD); })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, []);

  const stokStr = data?.stok
    ? new Date(data.stok.updatedAt).toLocaleString("tr-TR", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit", timeZone: "Europe/Istanbul" })
    : "";

  return (
    <div className="dash">
      <div className="kpi-row">
        <StatTile label="Stok Kalemi" value={data?.stok ? nf(data.stok.kalem) : "—"} ratio={1} ratioLabel={stokStr ? `Güncelleme: ${stokStr}` : "Güncel stok"} icon="package" href="/portal" tone="blue" loading={loading} />
        <StatTile label="Çerçeve Profili" value={data ? nf(data.katalog.profil) : "—"} ratio={1} ratioLabel={data ? `${data.katalog.seri.length} seri` : ""} icon="layers" href="/portal/fiyat-listesi" tone="brand" loading={loading} />
        <StatTile label="Teknik Malzeme" value={data ? nf(data.katalog.teknik) : "—"} ratio={1} ratioLabel="Fiyat listesinde" icon="box" href="/portal/fiyat-listesi" tone="green" loading={loading} />
        <a href="tel:+908503057545" className="stat card" title="Sipariş Hattı">
          <div className="stat-ring tone-amber" aria-hidden>
            <svg width="68" height="68" viewBox="0 0 68 68" style={{ width: "100%", height: "100%" }}>
              <circle cx="34" cy="34" r="26" className="stat-ring-track" />
              <circle cx="34" cy="34" r="26" className="stat-ring-fill" />
            </svg>
            <span className="stat-ring-icon"><Icon name="phone" size={18} /></span>
          </div>
          <div className="stat-body" style={PHONE_BODY}>
            <span className="stat-label">Sipariş Hattı</span>
            <span className="stat-value" style={PHONE_VALUE}>0850&nbsp;305 75&nbsp;45</span>
            <span className="stat-foot">
              <span className="stat-ratio" style={PHONE_LINE}>Ankara 0312&nbsp;495&nbsp;75&nbsp;45</span>
              <span className="stat-ratio" style={PHONE_LINE}>İstanbul 0212&nbsp;675&nbsp;27&nbsp;50</span>
            </span>
          </div>
        </a>
      </div>
      <div className="dash-grid">
        <section className="card span-12">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="zap" size={18} /></span>
            <div>
              <h2>Hızlı Erişim</h2>
              <span className="card-head-sub">Stok, fiyat listesi ve kataloglar</span>
            </div>
          </div>
          <QuickActions tiles={tiles} />
        </section>
      </div>
    </div>
  );
}
