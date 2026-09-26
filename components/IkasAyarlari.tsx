"use client";

// Sahiplere/üretim yetkililerine özel: ikas (olgacerceve.com) bağlantı durumu,
// bağlantı sınaması, son siparişlerde not çözümleme önizlemesi, webhook kurulumu
// ve sipariş numarasıyla tek seferlik içe aktarma.

import { useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Durum { ikas: boolean; webhookKey: boolean; db: boolean; webhookAdresi: string }
interface Satir { ad?: string; sku?: string; adet?: number }
interface Siparis {
  id: string; no?: string | null; tarih?: string | null; durum?: string | null; paket?: string | null; kanal?: string | null;
  musteri?: string; tutar?: number | null; satirlar: Satir[]; not?: string | null;
  cozum: { kalemler: { sku: string; adet: number; ozet: string; retail?: string }[]; eksik: string[]; notVar: boolean };
}
interface Yanit { ok: boolean; durum?: Durum; error?: string; me?: { id: string } | null; toplamSiparis?: number; siparisler?: Siparis[]; webhooklar?: unknown }

export default function IkasAyarlari() {
  const [y, setY] = useState<Yanit | null>(null);
  const [yukleniyor, setYukleniyor] = useState(true);
  const [islem, setIslem] = useState("");
  const [mesaj, setMesaj] = useState("");
  const [hata, setHata] = useState("");
  const [siparisNo, setSiparisNo] = useState("");

  async function yukle() {
    setYukleniyor(true); setHata("");
    try {
      const r = await fetch("/api/ikas/test", { cache: "no-store" });
      setY((await r.json()) as Yanit);
    } catch { setHata("Sunucuya ulaşılamadı."); }
    finally { setYukleniyor(false); }
  }
  useEffect(() => { yukle(); }, []);

  async function gonder(govde: Record<string, unknown>, etiket: string) {
    setIslem(etiket); setMesaj(""); setHata("");
    try {
      const r = await fetch("/api/ikas/test", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(govde) });
      const d = await r.json();
      if (!r.ok || !d.ok) setHata(d.error || "İşlem başarısız.");
      else setMesaj(etiket === "siparis-al" ? `Sipariş takvime alındı (${d.sonuc?.yeni ? "yeni iş" : "güncellendi"}${d.sonuc?.foy ? ", föy üretildi" : ""}).` : etiket === "webhook-kur" ? "Webhook kuruldu." : "Webhook silindi.");
      await yukle();
    } catch { setHata("Sunucuya ulaşılamadı."); }
    finally { setIslem(""); }
  }

  const d = y?.durum;
  const Satir = ({ ok, baslik, aciklama }: { ok: boolean; baslik: string; aciklama: string }) => (
    <li className="row" style={{ padding: "10px 0", borderBottom: "1px solid var(--border)", alignItems: "flex-start", flexWrap: "nowrap" }}>
      <span className={`badge ${ok ? "ok" : "warn"}`} style={{ marginTop: 2 }}>
        <Icon name={ok ? "check-circle" : "alert"} size={13} /> {ok ? "Hazır" : "Eksik"}
      </span>
      <div style={{ minWidth: 0 }}>
        <strong>{baslik}</strong>
        <div className="muted" style={{ fontSize: 12.5 }}>{aciklama}</div>
      </div>
    </li>
  );

  return (
    <div className="card">
      <div className="card-head">
        <span className="card-head-icon"><Icon name="calendar" size={18} /></span>
        <div>
          <h2>ikas — Online Siparişler → Üretim Takvimi</h2>
          <span className="card-head-sub">olgacerceve.com siparişleri özel uygulama API&apos;siyle takvime düşer; webhook + açılış senkronu</span>
        </div>
        <div className="card-head-actions">
          <button type="button" className={`btn small secondary ${yukleniyor ? "spin" : ""}`} onClick={yukle} disabled={yukleniyor} title="Yenile">
            <Icon name="refresh" size={14} /> Sına
          </button>
        </div>
      </div>

      {hata && <div className="notice err" style={{ marginBottom: 10 }}>{hata}</div>}
      {mesaj && <div className="notice ok" style={{ marginBottom: 10 }}>{mesaj}</div>}

      <ul style={{ listStyle: "none", padding: 0, margin: 0 }}>
        <Satir ok={Boolean(d?.ikas)} baslik="Özel uygulama anahtarları (IKAS_CLIENT_ID / IKAS_CLIENT_SECRET)" aciklama="ikas paneli → Uygulamalar → Uygulamalarım → Özel uygulama oluştur (kapsam: read_orders). IKAS_STORE = mağaza adı (olgacerceve)." />
        <Satir ok={Boolean(d?.db)} baslik="Veri tabanı (DATABASE_URL)" aciklama="Üretim takvimi işleri Postgres'te tutulur (mesajlarla aynı bağlantı)." />
        <Satir ok={Boolean(d?.webhookKey)} baslik="Webhook anahtarı (IKAS_WEBHOOK_KEY)" aciklama={d?.webhookAdresi ? `Webhook adresi: ${d.webhookAdresi}` : "Rastgele uzun bir metin; webhook adresinin sonuna ?k=… olarak eklenir. İmza (client secret HMAC) de doğrulanır."} />
      </ul>

      {y && !y.ok && y.error && <div className="notice warn" style={{ marginTop: 12 }}>{y.error}</div>}

      {y?.ok && (
        <>
          <div className="row" style={{ gap: 8, marginTop: 12, flexWrap: "wrap", alignItems: "center" }}>
            <span className="badge ok"><Icon name="check-circle" size={13} /> Bağlantı tamam · uygulama {y.me?.id?.slice(0, 8)}…</span>
            <span className="badge">{y.toplamSiparis ?? 0} sipariş</span>
            <span className="spacer" />
            <button type="button" className="btn small" onClick={() => gonder({ islem: "webhook-kur" }, "webhook-kur")} disabled={!!islem || !d?.webhookKey} title={d?.webhookKey ? "store/order/created + updated" : "Önce IKAS_WEBHOOK_KEY tanımlayın"}>
              <Icon name="zap" size={14} /> Webhook&apos;u kur
            </button>
            <button type="button" className="btn small secondary" onClick={() => gonder({ islem: "webhook-sil" }, "webhook-sil")} disabled={!!islem}>
              Webhook&apos;u sil
            </button>
          </div>
          <div className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>
            Kayıtlı webhook&apos;lar: {Array.isArray(y.webhooklar) && y.webhooklar.length
              ? (y.webhooklar as { scope: string; endpoint: string }[]).map((w) => `${w.scope} → ${w.endpoint.replace(/k=[^&]+/, "k=•••")}`).join(" · ")
              : "yok (webhook olmadan da takvim her açılışta ve sabah cron'unda senkronlar)"}
          </div>

          <div className="row" style={{ gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <input value={siparisNo} onChange={(e) => setSiparisNo(e.target.value)} placeholder="Sipariş no (örn. 1463)" style={{ maxWidth: 220 }} />
            <button type="button" className="btn small secondary" onClick={() => gonder({ islem: "siparis-al", orderNumber: siparisNo.trim() }, "siparis-al")} disabled={!!islem || !siparisNo.trim()}>
              <Icon name="download" size={14} /> Siparişi takvime al
            </button>
          </div>

          {y.siparisler && y.siparisler.length > 0 && (
            <div style={{ marginTop: 14 }}>
              <strong style={{ fontSize: 13.5 }}>Son siparişler ve not çözümü</strong>
              <div className="muted" style={{ fontSize: 12.5, margin: "4px 0 8px" }}>
                ikas zaman çizelgesindeki &quot;Cerceve Hesaplayici&quot; notu API&apos;den okunamaz. Kalemler sipariş notu / satır seçenekleri / varyant adından çözülür;
                tam ölçü-paspartu-fiyat detayı için hesaplayıcı sunucusu aynı metni <code>/api/ikas/not?k=…</code> ucuna da göndermelidir.
              </div>
              <div className="table-wrap">
                <table>
                  <thead><tr><th>No</th><th>Müşteri</th><th>Durum</th><th>Satırlar</th><th>Çözüm</th></tr></thead>
                  <tbody>
                    {y.siparisler.map((s) => (
                      <tr key={s.id}>
                        <td>#{s.no}<div className="muted" style={{ fontSize: 11.5 }}>{s.tarih ? new Date(s.tarih).toLocaleDateString("tr-TR") : ""}</div></td>
                        <td>{s.musteri}<div className="muted" style={{ fontSize: 11.5 }}>₺{(Number(s.tutar) || 0).toLocaleString("tr-TR")}</div></td>
                        <td><span className="badge info">{s.durum}</span><div className="muted" style={{ fontSize: 11.5 }}>{s.paket || ""}</div></td>
                        <td style={{ fontSize: 12.5 }}>{s.satirlar.map((x, i) => <div key={i}>{x.adet}× {x.ad} <span className="muted">({x.sku})</span></div>)}</td>
                        <td style={{ fontSize: 12.5 }}>
                          {s.cozum.kalemler.map((k, i) => (
                            <div key={i}>{k.adet}× <b>{k.sku}</b> — {k.ozet} {k.retail ? <span className="badge ok" style={{ fontSize: 10.5 }}>föy</span> : <span className="badge warn" style={{ fontSize: 10.5 }}>ölçü yok</span>}</div>
                          ))}
                          {s.cozum.eksik.length > 0 && <div className="muted">Çözülemeyen: {s.cozum.eksik.join(", ")}</div>}
                          <div className="muted" style={{ fontSize: 11.5 }}>{s.cozum.notVar ? "Not metni var" : "Not metni yok"}</div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}
