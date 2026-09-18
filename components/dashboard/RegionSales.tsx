"use client";

// Bölge cirosu: Ankara / İstanbul / Taşra / kayıtsız müşteri — bu ay, önceki
// ayla kıyas ve toplam içindeki pay. Bölge müşteri kartından gelir; siparişi
// kim aldıysa alsın. Kayıtsız: sipariş deftere girilmemiş bir müşteriye
// yazılmış (bölgesi bilinmiyor) — kutuya tıklayınca müşteri defteri açılır.

import Link from "next/link";
import Icon from "@/components/shell/Icon";
import type { StaffDashboard } from "@/lib/dashboard";

const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });
const nf = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");

export default function RegionSales({ data, loading }: { data?: StaffDashboard["bolge"]; loading: boolean }) {
  if (!data && !loading) return null;
  return (
    <section className="card span-12">
      <div className="card-head">
        <span className="card-head-icon"><Icon name="map-pin" size={18} /></span>
        <div>
          <h2>Bölge Cirosu{data ? ` · ${data.ayLabel}` : ""}</h2>
          <span className="card-head-sub">
            Müşteri kartındaki bölgeye göre · toptan + perakende, iptaller hariç{data ? ` · toplam ${tl(data.toplam)}` : ""}
          </span>
        </div>
        <span className="spacer" />
        <Link href="/musteriler" className="btn small secondary">Müşteriler</Link>
      </div>
      {!data ? (
        <div className="rg-grid">{[0, 1, 2, 3].map((i) => <div key={i} className="skeleton" style={{ height: 104, borderRadius: 11 }} />)}</div>
      ) : (
        <div className="rg-grid">
          {data.kalemler.map((k) => {
            const fark = k.ciro - k.oncekiCiro;
            const delta = k.oncekiCiro > 0
              ? `${fark >= 0 ? "+" : "−"}%${Math.round((Math.abs(fark) / k.oncekiCiro) * 100)} önceki aya göre`
              : k.ciro > 0 ? "önceki ay veri yok" : "";
            const href = k.id === "kayitsiz" ? "/musteriler" : `/musteriler?bolge=${k.id}`;
            return (
              <Link key={k.id} href={href} className={`rg-tile ${k.id}`} title={k.id === "kayitsiz" ? "Müşteri deftere girilmemiş siparişler — müşteriyi kaydedip siparişte listeden seçin" : `${k.label} müşterilerinin siparişleri`}>
                <div className="rg-head">
                  <span className="rg-label">{k.label}</span>
                  <span className="rg-pay">%{Math.round(k.pay * 100)}</span>
                </div>
                <div className="rg-val">{tl(k.ciro)}</div>
                <div className="rg-bar"><i style={{ width: `${Math.round(k.pay * 100)}%` }} /></div>
                <div className="rg-sub">
                  <span>{nf(k.adet)} sipariş · {tl(k.toptanCiro)} toptan · {tl(k.perakendeCiro)} perakende</span>
                </div>
                {delta && <div className={`rg-delta ${fark > 0 ? "up" : fark < 0 ? "down" : ""}`}>{delta}</div>}
              </Link>
            );
          })}
        </div>
      )}
      {data && data.musteriSayisi === 0 && (
        <p className="muted" style={{ fontSize: 12.5, marginTop: 10 }}>Müşteri defteri boş; bölge ayrımı için müşterileri kaydedip siparişte listeden seçin.</p>
      )}
    </section>
  );
}
