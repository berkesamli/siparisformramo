"use client";

// Müşteri / bayi ana sayfası: hızlı erişim kutuları + katalog/stok özeti.

import { useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";
import type { CustomerDashboard as CD } from "@/lib/dashboard";
import StatTile from "./StatTile";
import QuickActions, { type QuickTile } from "./QuickActions";

const nf = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");

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
        <StatTile label="Sipariş Hattı" value="0850 305 75 45" ratio={1} ratioLabel="Ankara 0312 495 75 45 · İstanbul 0212 675 27 50" icon="phone" tone="amber" />
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
