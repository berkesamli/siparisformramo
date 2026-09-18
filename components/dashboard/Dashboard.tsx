"use client";

// Çalışan gösterge paneli: KPI halkaları, 14 günlük akış, durum halkası,
// son siparişler, uyarılar. Veri /api/dashboard'dan gelir; yenilemede
// önceki çizim solgun kalır (iskelet flaşı yok).

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import type { StaffDashboard } from "@/lib/dashboard";
import StatTile from "./StatTile";
import RegionSales from "./RegionSales";
import OrderFlowChart from "./OrderFlowChart";
import StatusDonut from "./StatusDonut";
import { RecentWholesale, RecentRetail } from "./RecentOrders";
import AlertsCard from "./AlertsCard";
import QuickActions, { type QuickTile } from "./QuickActions";
import MySalesCard from "@/components/sales/MySalesCard";

const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });
const nf = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");

const BOS: StaffDashboard = {
  role: "staff", blob: false, today: "", hesaplandi: "",
  lite: { bugun: 0, bugunToptan: 0, bugunPerakende: 0, dun: 0, acik: 0, perakendeAcik: 0, kontrolsuz: 0, ayAdet: 0, blob: false },
  kpi: { bugun: { toptan: 0, perakende: 0, toplam: 0, dun: 0 }, acik: { toptan: 0, perakende: 0 }, kontrolsuz: 0, ay: { toptanAdet: 0, perakendeAdet: 0, toptanCiro: 0, perakendeCiro: 0 }, gun14: { toptan: 0, perakende: 0, toptanCiro: 0, perakendeCiro: 0 } },
  seri14: [], durum14: { toptan: { olusturuldu: 0, hazirlaniyor: 0, tamamlandi: 0, iptal: 0 }, perakende: { Beklemede: 0, "Hazırlanıyor": 0, "Hazır": 0, "Teslim Edildi": 0, "İptal": 0 } },
  sonToptan: [], sonPerakende: [], uyarilar: [], stok: null, kur: null,
};

export default function Dashboard({
  finance,
  tiles,
  chips,
}: {
  finance: boolean;
  tiles: QuickTile[];
  chips: { label: string; href: string; icon: "users" | "message" | "upload" | "tag" | "book" | "dollar" | "bar-chart" }[];
}) {
  const [data, setData] = useState<StaffDashboard | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [mode, setMode] = useState<"adet" | "ciro">("adet");

  const load = useCallback(async () => {
    setLoading(true);
    setError("");
    try {
      const r = await fetch("/api/dashboard");
      const d = await r.json();
      if (d.ok) setData(d.data as StaffDashboard);
      else setError(d.error || "Veriler alınamadı.");
    } catch {
      setError("Sunucuya ulaşılamadı.");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { load(); }, [load]);

  const d = data || BOS;
  const k = d.kpi;
  // Ay ilerlemesi sunucunun `today` değerinden türetilir; render içinde new Date()
  // kullanılmaz (sunucu/istemci saat dilimi farkı hidrasyon uyuşmazlığı yaratmasın).
  const [tYil, tAy, tGun] = d.today.split("-").map(Number);
  const ayDay = tGun || 0;
  const ayGun = tYil && tAy ? new Date(Date.UTC(tYil, tAy, 0)).getUTCDate() : 30;
  const tamamlanan14 = d.durum14.toptan.tamamlandi + d.durum14.perakende["Teslim Edildi"] + d.durum14.perakende["Hazır"];
  const acikToplam = k.acik.toptan + k.acik.perakende;
  const bugunDelta = k.bugun.toplam - k.bugun.dun;
  const kontrolOran = k.gun14.toptan > 0 ? 1 - Math.min(1, k.kontrolsuz / Math.max(1, k.gun14.toptan)) : 1;

  const stokStr = d.stok
    ? new Date(d.stok.updatedAt).toLocaleString("tr-TR", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit", timeZone: "Europe/Istanbul" })
    : "—";

  return (
    <div className="dash">
      {error && (
        <div className="notice err row">
          <span style={{ flex: 1 }}>{error}</span>
          <button type="button" className="btn small secondary" onClick={load}>Tekrar dene</button>
        </div>
      )}

      <div className="kpi-row">
        <StatTile
          label="Bugünkü Siparişler"
          value={nf(k.bugun.toplam)}
          ratio={k.bugun.toplam + k.bugun.dun > 0 ? k.bugun.toplam / (k.bugun.toplam + k.bugun.dun) : 0}
          delta={data ? `${bugunDelta >= 0 ? "+" : ""}${bugunDelta} düne göre` : undefined}
          deltaGood={bugunDelta > 0 ? true : bugunDelta < 0 ? false : null}
          ratioLabel={`${k.bugun.toptan} toptan · ${k.bugun.perakende} perakende`}
          icon="inbox"
          href="/panel/siparisler"
          tone="blue"
          loading={loading && !data}
        />
        <StatTile
          label="Açık Siparişler"
          value={nf(acikToplam)}
          ratio={acikToplam + tamamlanan14 > 0 ? tamamlanan14 / (acikToplam + tamamlanan14) : 0}
          ratioLabel={`${k.acik.toptan} toptan · ${k.acik.perakende} perakende`}
          delta={data ? `${tamamlanan14} tamamlandı (14 gün)` : undefined}
          icon="activity"
          href="/panel/siparisler"
          tone="amber"
          loading={loading && !data}
        />
        <StatTile
          label="Kontrol Bekleyen"
          value={nf(k.kontrolsuz)}
          ratio={kontrolOran}
          ratioLabel={`%${Math.round(kontrolOran * 100)} kontrol edildi`}
          delta={data ? "son 7 gün" : undefined}
          deltaGood={k.kontrolsuz === 0 ? true : null}
          icon="eye"
          href="/panel/siparisler"
          tone={k.kontrolsuz > 0 ? "red" : "green"}
          loading={loading && !data}
        />
        {finance ? (
          <StatTile
            label="Bu Ay Ciro"
            value={tl(k.ay.toptanCiro + k.ay.perakendeCiro)}
            ratio={ayDay / ayGun}
            ratioLabel={`${tl(k.ay.toptanCiro)} toptan · ${tl(k.ay.perakendeCiro)} perakende`}
            delta={data ? `ayın %${Math.round((ayDay / ayGun) * 100)}'i geçti` : undefined}
            icon="trending-up"
            href="/panel/raporlar"
            tone="green"
            loading={loading && !data}
          />
        ) : (
          <StatTile
            label="Bu Ay Sipariş"
            value={nf(k.ay.toptanAdet + k.ay.perakendeAdet)}
            ratio={ayDay / ayGun}
            ratioLabel={`${k.ay.toptanAdet} toptan · ${k.ay.perakendeAdet} perakende`}
            delta={data ? `ayın %${Math.round((ayDay / ayGun) * 100)}'i geçti` : undefined}
            icon="calendar"
            href="/panel/siparisler"
            tone="green"
            loading={loading && !data}
          />
        )}
      </div>

      <div className="dash-grid">
        {/* Bölge cirosu — finans yetkisi (sahip): Ankara / İstanbul / Taşra / kayıtsız */}
        {(d.bolge || (finance && loading && !data)) && <RegionSales data={d.bolge} loading={loading && !data} />}
        <section className="card span-8">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="bar-chart" size={18} /></span>
            <div>
              <h2>Sipariş Akışı</h2>
              <span className="card-head-sub">Son 14 gün · günlük {mode === "adet" ? "adet" : "ciro"}</span>
            </div>
            <span className="spacer" />
            {finance && (
              <div className="seg">
                <button type="button" className={mode === "adet" ? "active" : ""} onClick={() => setMode("adet")}>Adet</button>
                <button type="button" className={mode === "ciro" ? "active" : ""} onClick={() => setMode("ciro")}>Ciro</button>
              </div>
            )}
            <button type="button" className="btn icon ghost small" onClick={load} title="Yenile" aria-label="Yenile">
              <Icon name="refresh" size={16} className={loading ? "spin" : undefined} />
            </button>
          </div>
          <OrderFlowChart seri={d.seri14} mode={mode} loading={loading} />
        </section>

        <section className="card span-4">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="pie" size={18} /></span>
            <div>
              <h2>Durum Dağılımı</h2>
              <span className="card-head-sub">Son 14 gün</span>
            </div>
          </div>
          <StatusDonut toptan={d.durum14.toptan} perakende={d.durum14.perakende} loading={loading} />
        </section>

        <section className="card span-7 pad-0">
          <div className="card-head" style={{ margin: 0 }}>
            <span className="card-head-icon"><Icon name="list" size={18} /></span>
            <div>
              <h2>Son Toptan Siparişler</h2>
              <span className="card-head-sub">{d.blob ? "En yeni 8 kayıt" : "Depo bağlı değil"}</span>
            </div>
            <span className="spacer" />
            <Link href="/panel/siparisler" className="btn small secondary">Tümü</Link>
          </div>
          <RecentWholesale orders={d.sonToptan} loading={loading} blob={d.blob} />
        </section>

        <section className="card span-5 pad-0">
          <div className="card-head" style={{ margin: 0 }}>
            <span className="card-head-icon"><Icon name="alert" size={18} /></span>
            <div>
              <h2>Uyarılar</h2>
              <span className="card-head-sub">{d.uyarilar.length ? `${d.uyarilar.length} konu` : "Takip gerektiren konu yok"}</span>
            </div>
          </div>
          <AlertsCard uyarilar={d.uyarilar} loading={loading} />
          <div className="dash-meta">
            <span><Icon name="package" size={14} /> Stok: {d.stok ? `${nf(d.stok.kalem)} kalem · ${stokStr}` : "—"}</span>
            <span><Icon name="dollar" size={14} /> Kur: {d.kur ? `$ ${d.kur.rate.toLocaleString("tr-TR")} · € ${d.kur.euroRate.toLocaleString("tr-TR")}` : "bugün girilmedi"}</span>
          </div>
        </section>

        {/* Çalışanın kendi satış özeti — kendi <section className="card span-5"> kabuğunu çizer */}
        <MySalesCard />

        <section className="card span-7 pad-0">
          <div className="card-head" style={{ margin: 0 }}>
            <span className="card-head-icon"><Icon name="frame" size={18} /></span>
            <div>
              <h2>Son Perakende Siparişler</h2>
              <span className="card-head-sub">Online çerçeve</span>
            </div>
            <span className="spacer" />
            <Link href="/panel/perakende/siparisler" className="btn small secondary">Tümü</Link>
          </div>
          <RecentRetail orders={d.sonPerakende} loading={loading} blob={d.blob} />
        </section>

        <section className="card span-12">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="zap" size={18} /></span>
            <div>
              <h2>Hızlı İşlemler</h2>
              <span className="card-head-sub">En sık kullanılan ekranlar</span>
            </div>
          </div>
          <QuickActions tiles={tiles} chips={chips} />
        </section>
      </div>
    </div>
  );
}
