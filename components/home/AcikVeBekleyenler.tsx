"use client";

// Ana sayfa listeleri:
// - Açık Toptan Siparişler: en yeni önce; durum satırdan değiştirilir
//   (seçici ya da tek dokunuşla "sonraki adım"), detaya gitmeden.
// - Beni Bekleyenler: son 7 günde merkez kontrolü yapılmamış siparişler;
//   "Kontrol edildi" ile listeden düşer.
// Veri /api/orders/acik; güncellemeler /api/orders/one PATCH (Siparişler
// sayfasıyla aynı uç). İyimser güncelleme: satır anında değişir, hata olursa
// liste yeniden yüklenir.

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import { STATUS_LABELS, type OrderIndexEntry, type OrderStatus } from "@/lib/orders";

interface Veri {
  ok: boolean;
  blob: boolean;
  today: string;
  acik: OrderIndexEntry[];
  toplamAcik: number;
  kontrolsuz: OrderIndexEntry[];
  error?: string;
}

const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });
const saat = (iso: string) => new Date(iso).toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit", timeZone: "Europe/Istanbul" });
const gun = (k: string) => { const [, m, d] = k.split("-"); return `${Number(d)}.${m}`; };
const detay = (o: OrderIndexEntry) => `/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`;

// Tek dokunuşla ilerletme: Oluşturuldu → Hazırlanıyor → Tamamlandı (yarım → tamamlandı)
const SONRAKI: Partial<Record<OrderStatus, OrderStatus>> = { olusturuldu: "hazirlaniyor", hazirlaniyor: "tamamlandi", yarim: "tamamlandi" };

export default function AcikVeBekleyenler() {
  const [veri, setVeri] = useState<Veri | null>(null);
  const [loading, setLoading] = useState(true);
  const [hata, setHata] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    setHata("");
    try {
      const r = await fetch("/api/orders/acik");
      const d = (await r.json()) as Veri;
      if (d.ok) setVeri(d);
      else setHata(d.error || "Veriler alınamadı.");
    } catch {
      setHata("Sunucuya ulaşılamadı.");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { load(); }, [load]);

  async function patch(o: OrderIndexEntry, body: Record<string, unknown>): Promise<boolean> {
    const r = await fetch(`/api/orders/one?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }).catch(() => null);
    return Boolean(r && r.ok);
  }

  async function durumDegistir(o: OrderIndexEntry, status: OrderStatus) {
    if (status === o.status) return;
    if (status === "iptal") {
      const onay = confirm(
        `${o.orderId} — ${o.customer || "müşteri"} siparişi iptal edilsin mi?\n\n` +
          "İptal edilen sipariş silinmez ama ciroya, raporlara ve müşteri bakiyesine dahil edilmez."
      );
      if (!onay) return;
    }
    setVeri((v) => v && { ...v, acik: v.acik.map((x) => (x.orderId === o.orderId ? { ...x, status } : x)) });
    const ok = await patch(o, { status });
    if (!ok) { setHata("Durum güncellenemedi, liste yenileniyor."); load(); return; }
    if (status === "tamamlandi" || status === "iptal") {
      // Kapanan sipariş kısa bir an solgun görünsün, sonra listeden düşsün
      setTimeout(() => {
        setVeri((v) => v && {
          ...v,
          acik: v.acik.filter((x) => x.orderId !== o.orderId),
          toplamAcik: Math.max(0, v.toplamAcik - 1),
        });
      }, 700);
    }
  }

  async function kontrolEdildi(o: OrderIndexEntry) {
    setVeri((v) => v && { ...v, kontrolsuz: v.kontrolsuz.filter((x) => x.orderId !== o.orderId) });
    const ok = await patch(o, { kontrol: true });
    if (!ok) { setHata("Kontrol işareti kaydedilemedi, liste yenileniyor."); load(); }
  }

  const acik = veri?.acik || [];
  const bekleyen = veri?.kontrolsuz || [];
  const bugunAcik = veri ? acik.filter((o) => o.dateKey === veri.today).length : 0;

  return (
    <div className="dash-grid home-lists">
      {hata && (
        <div className="notice err row span-12" style={{ margin: 0 }}>
          <span style={{ flex: 1 }}>{hata}</span>
          <button type="button" className="btn small secondary" onClick={load}>Tekrar dene</button>
        </div>
      )}

      <section className="card span-7 pad-0">
        <div className="card-head" style={{ margin: 0 }}>
          <span className="card-head-icon"><Icon name="activity" size={18} /></span>
          <div>
            <h2>Açık Toptan Siparişler</h2>
            <span className="card-head-sub">
              {veri ? `${veri.toplamAcik} açık · ${bugunAcik} bugün girildi` : "Yükleniyor…"}
            </span>
          </div>
          <span className="spacer" />
          <button type="button" className="btn icon ghost small" onClick={load} title="Yenile" aria-label="Yenile">
            <Icon name="refresh" size={16} className={loading ? "spin" : undefined} />
          </button>
          <Link href="/panel/siparisler" className="btn small secondary">Tümü</Link>
        </div>
        {!loading && acik.length === 0 ? (
          <div className="empty">
            <div className="empty-icon" style={{ color: "var(--success)", background: "var(--success-soft)" }}><Icon name="check-circle" size={22} /></div>
            <strong>Açık sipariş yok</strong>
            {veri?.blob === false ? "Kalıcı depolama bağlandığında siparişler burada görünür." : "Yeni sipariş girildiğinde burada listelenir."}
          </div>
        ) : (
          <ul className={`ao-list ${loading ? "loading-dim" : ""}`}>
            {acik.map((o) => {
              const sonraki = SONRAKI[o.status];
              return (
                <li key={o.orderId} className={`ao-row ${o.status}`}>
                  <Link href={detay(o)} className="ao-main" title="Sipariş detayı">
                    <span className="ao-title">{o.customer || "—"}</span>
                    <span className="ao-sub">
                      {o.orderId} · {gun(o.dateKey)} {saat(o.createdAt)} · {o.employee}
                      {veri && o.dateKey === veri.today ? " · bugün" : ""}
                    </span>
                  </Link>
                  <span className="ao-amt num">{tl(o.net)}</span>
                  <select
                    className={`status-select ${o.status}`}
                    value={o.status}
                    onChange={(e) => durumDegistir(o, e.target.value as OrderStatus)}
                    aria-label={`${o.orderId} durumu`}
                  >
                    {Object.entries(STATUS_LABELS).map(([k, v]) => (
                      <option key={k} value={k}>{v}</option>
                    ))}
                  </select>
                  {sonraki && (
                    <button
                      type="button"
                      className="btn xs secondary ao-next"
                      onClick={() => durumDegistir(o, sonraki)}
                      title={`Durumu "${STATUS_LABELS[sonraki]}" yap`}
                    >
                      <Icon name="chevron-right" size={14} /> {STATUS_LABELS[sonraki]}
                    </button>
                  )}
                </li>
              );
            })}
          </ul>
        )}
      </section>

      <section className="card span-5 pad-0">
        <div className="card-head" style={{ margin: 0 }}>
          <span className="card-head-icon"><Icon name="eye" size={18} /></span>
          <div>
            <h2>Beni Bekleyenler</h2>
            <span className="card-head-sub">
              {veri ? (bekleyen.length ? `${bekleyen.length} sipariş kontrol bekliyor` : "Kontrol bekleyen yok") : "Yükleniyor…"}
            </span>
          </div>
          <span className="spacer" />
          <Link href="/panel/siparisler" className="btn small secondary">Siparişler</Link>
        </div>
        {!loading && bekleyen.length === 0 ? (
          <div className="empty" style={{ padding: "26px 12px" }}>
            <div className="empty-icon" style={{ color: "var(--success)", background: "var(--success-soft)" }}><Icon name="check-circle" size={22} /></div>
            <strong>Her şey yolunda</strong>
            Son 7 günde kontrol bekleyen sipariş yok.
          </div>
        ) : (
          <ul className={`bb-list ${loading ? "loading-dim" : ""}`}>
            {bekleyen.map((o) => (
              <li key={o.orderId} className="bb-row">
                <Link href={detay(o)} className="bb-main" title="Sipariş detayı">
                  <span className="bb-title">{o.customer || "—"}</span>
                  <span className="bb-sub">{o.orderId} · {gun(o.dateKey)} · {tl(o.net)} · {o.employee}</span>
                </Link>
                <button type="button" className="btn small secondary" onClick={() => kontrolEdildi(o)} title="Merkez kontrolü yapıldı olarak işaretle">
                  <Icon name="check-circle" size={15} /> Kontrol edildi
                </button>
              </li>
            ))}
          </ul>
        )}
      </section>
    </div>
  );
}
