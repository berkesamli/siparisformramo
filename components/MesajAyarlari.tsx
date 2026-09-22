"use client";

// Sahiplere özel: Mesajlar (gelen kutusu) kurulum durumu, canlı bağlantı testleri, e-postaları şimdi çekme.

import { useEffect, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";

interface Durum {
  db: { kurulu: boolean; test: { ok: boolean; konusma: number; mesaj: number; hata?: string; sunucu?: string } | null };
  gmail: { kurulu: boolean; hesaplar: string[]; test: { adres: string; ok: boolean; inbox?: number; okunmamis?: number; hata?: string }[] | null };
  whatsapp: {
    kurulu: boolean; sablonlar: string[]; test: { ok: boolean; numara?: string; ad?: string; kalite?: string; hata?: string } | null;
    webhook: { url: string; verifyToken: boolean; appSecret: boolean; sonOlay: Iz | null; wabaId: boolean; abonelik: Abonelik | null };
  };
  instagram: { kurulu: boolean; test: { ok: boolean; ad?: string; hata?: string } | null; webhook: { url: string; sonOlay: Iz | null } };
  taslak: boolean;
  gorebilenler: string[];
}
interface SenkSonuc { hesap: string; yeni: number; hata?: string; atlandi?: boolean }
interface Iz { at: string; tur: string; ozet: string }
interface Abonelik { ok: boolean; wabaId?: string; abone?: boolean; alanlar?: string[]; hata?: string }

const nekadar = (iso: string) => {
  const dk = Math.round((Date.now() - new Date(iso).getTime()) / 60_000);
  if (dk < 1) return "az önce"; if (dk < 60) return `${dk} dk önce`; if (dk < 48 * 60) return `${Math.round(dk / 60)} saat önce`;
  return new Date(iso).toLocaleString("tr-TR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
};

function Satir({ ok, baslik, detay, uyari }: { ok: boolean | null; baslik: string; detay?: React.ReactNode; uyari?: boolean }) {
  return (
    <div className="ma-row">
      <span className={`badge ${ok === null ? "" : ok ? "ok" : uyari ? "warn" : "err"}`}>{ok === null ? "—" : ok ? "✓" : uyari ? "!" : "✗"}</span>
      <div><strong>{baslik}</strong>{detay && <div className="muted" style={{ fontSize: 12.5 }}>{detay}</div>}</div>
    </div>
  );
}

export default function MesajAyarlari() {
  const [d, setD] = useState<Durum | null>(null);
  const [hata, setHata] = useState("");
  const [testEdiyor, setTestEdiyor] = useState(false);
  const [cekiyor, setCekiyor] = useState(false);
  const [senk, setSenk] = useState<SenkSonuc[] | null>(null);
  const [senkHata, setSenkHata] = useState("");
  const [aboneOluyor, setAboneOluyor] = useState(false);
  const [aboneNotu, setAboneNotu] = useState("");

  async function yukle(test = false) {
    if (test) setTestEdiyor(true);
    try {
      const r = await fetch(`/api/mesaj/durum${test ? "?test=1" : ""}`, { cache: "no-store" });
      const j = await r.json();
      if (!j.ok) { setHata(j.error || "Durum alınamadı."); return; }
      setHata(""); setD(j);
    } catch { setHata("Sunucuya ulaşılamadı."); }
    finally { setTestEdiyor(false); }
  }
  useEffect(() => { void yukle(false); }, []);

  async function epostaCek() {
    setCekiyor(true); setSenk(null); setSenkHata("");
    try {
      const r = await fetch("/api/mesaj/durum", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ islem: "gmail-senk" }) });
      const j = await r.json();
      if (!j.ok) { setSenkHata(j.error || "Çekilemedi."); return; }
      setSenk(j.sonuc || []);
    } catch { setSenkHata("Sunucuya ulaşılamadı."); }
    finally { setCekiyor(false); }
  }

  async function aboneOl() {
    setAboneOluyor(true); setAboneNotu("");
    try {
      const r = await fetch("/api/mesaj/durum", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ islem: "wa-abone" }) });
      const j = await r.json();
      setAboneNotu(j.ok ? `Abone olundu (alanlar: ${(j.abonelik?.alanlar || []).join(", ") || "—"}). Şimdi numaraya bir mesaj atıp "Bağlantıları sına" ile son olayı kontrol edin.` : `Olmadı: ${j.error || "hata"}`);
      if (j.ok) void yukle(true);
    } catch { setAboneNotu("Sunucuya ulaşılamadı."); }
    finally { setAboneOluyor(false); }
  }

  const canli = Boolean(d?.db.test || d?.gmail.test || d?.whatsapp.test || d?.instagram.test);
  const wh = d?.whatsapp.webhook;
  const sonOlay = wh?.sonOlay;
  const webhookOk = sonOlay ? sonOlay.tur !== "imza-red" : null;
  const webhookDetay = !wh ? "" : sonOlay
    ? (sonOlay.tur === "imza-red"
        ? `Meta'dan olay geldi (${nekadar(sonOlay.at)}) ama reddedildi: ${sonOlay.ozet}`
        : `Meta'dan son olay ${nekadar(sonOlay.at)}: ${sonOlay.ozet}${sonOlay.tur !== "mesaj" ? " — mesaj içermiyor; gelen mesaj düşmüyorsa Meta'da \"messages\" alanına abonelik eksik olabilir." : ""}`)
    : `Meta'dan henüz hiç olay gelmedi. Meta uygulaması → WhatsApp → Configuration → Webhook: Callback URL ${wh.url}, Verify token = WHATSAPP_VERIFY_TOKEN${wh.verifyToken ? "" : " (Vercel'de tanımlı DEĞİL)"} → "Verify and save"; Webhook fields → messages → Subscribe.`;
  const abonelik = wh?.abonelik;

  return (
    <section className="card" style={{ marginTop: 16 }}>
      <div className="card-head">
        <span className="card-head-icon"><Icon name="inbox" size={20} /></span>
        <div>
          <h2>Mesajlar (Gelen Kutusu)</h2>
          <div className="card-head-sub">WhatsApp, Instagram ve Gmail bağlantıları; veri tabanı; kimler görüyor.</div>
        </div>
        <div className="card-head-actions">
          <Link href="/panel/mesajlar" className="btn secondary small"><Icon name="inbox" size={14} /> Gelen kutusunu aç</Link>
          <button type="button" className="btn small" onClick={() => void yukle(true)} disabled={testEdiyor}>
            <Icon name="zap" size={14} /> {testEdiyor ? "Sınanıyor…" : "Bağlantıları sına"}
          </button>
        </div>
      </div>

      {hata && <div className="notice err">{hata}</div>}
      {!d && !hata && <div className="skeleton" style={{ height: 80 }} />}
      {d && (
        <div className="ma-list">
          <Satir
            ok={d.db.test ? d.db.test.ok : d.db.kurulu ? null : false}
            baslik={`Veri tabanı ${d.db.kurulu ? "(bağlı)" : "(DATABASE_URL yok)"}`}
            detay={d.db.test ? (d.db.test.ok ? `${d.db.test.sunucu ? d.db.test.sunucu + " · " : ""}${d.db.test.konusma} konuşma, ${d.db.test.mesaj} mesaj` : d.db.test.hata) : d.db.kurulu ? "Sınamak için \"Bağlantıları sına\"." : "Vercel → Storage → Neon Postgres oluşturup projeye bağlayın."}
          />
          <Satir
            ok={d.gmail.test ? d.gmail.test.every((t) => t.ok) : d.gmail.kurulu ? null : false}
            uyari={!d.gmail.kurulu}
            baslik={`Gmail ${d.gmail.kurulu ? `(${d.gmail.hesaplar.length} hesap)` : "(GMAIL_HESAPLAR yok)"}`}
            detay={
              d.gmail.test
                ? <>{d.gmail.test.map((t) => <div key={t.adres}>{t.ok ? "✓" : "✗"} {t.adres}{t.ok ? ` — gelen kutusu ${t.inbox} e-posta, ${t.okunmamis} okunmamış` : ` — ${t.hata}`}</div>)}</>
                : d.gmail.kurulu ? d.gmail.hesaplar.join(", ") : "adres:uygulama-şifresi;adres2:şifre2 biçiminde girin."
            }
          />
          <Satir
            ok={d.whatsapp.test ? d.whatsapp.test.ok : d.whatsapp.kurulu ? null : false}
            uyari={!d.whatsapp.kurulu}
            baslik={`WhatsApp Cloud API ${d.whatsapp.kurulu ? "" : "(WHATSAPP_TOKEN / PHONE_ID yok)"}`}
            detay={
              d.whatsapp.test
                ? (d.whatsapp.test.ok ? `${d.whatsapp.test.numara || ""} ${d.whatsapp.test.ad ? "· " + d.whatsapp.test.ad : ""}${d.whatsapp.test.kalite ? " · kalite " + d.whatsapp.test.kalite : ""}` : d.whatsapp.test.hata)
                : d.whatsapp.sablonlar.length ? `Bizim başlattığımız mesaj şablonu: ${d.whatsapp.sablonlar.join(", ")}` : "Şablon tanımlı değil: yalnızca müşteri yazınca (24 saat içinde) yanıtlanır. Biz başlatmak için WHATSAPP_TEMPLATE_SERBEST."
            }
          />
          <Satir
            ok={webhookOk}
            uyari={sonOlay ? sonOlay.tur !== "mesaj" : true}
            baslik="WhatsApp webhook'u (gelen mesajlar)"
            detay={
              <>
                <div>{webhookDetay}</div>
                {wh && !wh.appSecret && <div><strong>Zorunlu:</strong> WHATSAPP_APP_SECRET tanımlı değil; güvenlik için Meta'dan gelen bütün olaylar reddedilir. Meta uygulaması → App settings → Basic → App secret değerini Vercel'e girin.</div>}
                {abonelik && (
                  <div>
                    {abonelik.ok
                      ? (abonelik.abone ? `Uygulama WhatsApp Business hesabına abone ✓${abonelik.alanlar?.length ? ` (alanlar: ${abonelik.alanlar.join(", ")})` : ""}${abonelik.alanlar && !abonelik.alanlar.includes("messages") ? " — \"messages\" alanı eksik!" : ""}` : "Uygulama WhatsApp Business hesabına ABONE DEĞİL → gelen mesajlar düşmez.")
                      : `Abonelik sorgulanamadı: ${abonelik.hata}`}
                  </div>
                )}
                {wh && !wh.wabaId && <div>Aboneliği buradan kontrol edip onarmak için Vercel'e WHATSAPP_WABA_ID (WhatsApp Business Account ID) ekleyin.</div>}
                {wh?.wabaId && (
                  <div style={{ marginTop: 6, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                    <button type="button" className="btn secondary xs" onClick={() => void aboneOl()} disabled={aboneOluyor}><Icon name="zap" size={12} /> {aboneOluyor ? "Abone olunuyor…" : "Webhook aboneliğini onar"}</button>
                    {aboneNotu && <span className={aboneNotu.startsWith("Olmadı") ? "" : ""} style={{ fontSize: 12.5 }}>{aboneNotu}</span>}
                  </div>
                )}
              </>
            }
          />
          <Satir
            ok={d.instagram.test ? d.instagram.test.ok : d.instagram.kurulu ? null : false}
            uyari={!d.instagram.kurulu}
            baslik={`Instagram ${d.instagram.kurulu ? "" : "(henüz bağlı değil)"}`}
            detay={<>{d.instagram.test ? (d.instagram.test.ok ? d.instagram.test.ad : d.instagram.test.hata) : d.instagram.kurulu ? "" : "Meta uygulamasında Instagram ürünü + instagram_manage_messages izni; INSTAGRAM_TOKEN, INSTAGRAM_PAGE_ID, INSTAGRAM_ACCOUNT_ID."}{d.instagram.webhook.sonOlay && <div>Webhook son olay {nekadar(d.instagram.webhook.sonOlay.at)}: {d.instagram.webhook.sonOlay.ozet}</div>}<div>Webhook adresi: {d.instagram.webhook.url}</div></>}
          />
          <Satir ok={d.taslak} uyari baslik="Yapay zekâ taslağı" detay={d.taslak ? "ANTHROPIC_API_KEY tanımlı; \"Taslak öner\" çalışır." : "ANTHROPIC_API_KEY yok; taslak düğmesi kapalı."} />
          <Satir ok={true} baslik="Kimler görüyor" detay={d.gorebilenler.join(", ")} />
          {!canli && <p className="muted" style={{ fontSize: 12.5, margin: "6px 0 0" }}>"Bağlantıları sına" veri tabanına, Gmail hesaplarına ve Meta'ya gerçekten bağlanıp sonucu gösterir (birkaç saniye sürer).</p>}
        </div>
      )}

      <div className="ma-actions">
        <button type="button" className="btn secondary small" onClick={() => void epostaCek()} disabled={cekiyor || !d?.gmail.kurulu || !d?.db.kurulu}>
          <Icon name="refresh" size={14} /> {cekiyor ? "Çekiliyor…" : "E-postaları şimdi çek"}
        </button>
        <span className="muted" style={{ fontSize: 12.5 }}>İlk çekimde son 7 günün e-postaları alınır; sonra gelen kutusu açıkken kendiliğinden tazelenir.</span>
      </div>
      {senkHata && <div className="notice err">{senkHata}</div>}
      {senk && (
        <div className={`notice ${senk.some((s) => s.hata) ? "warn" : "ok"}`}>
          {senk.map((s) => <div key={s.hesap}>{s.hesap}: {s.hata ? `hata — ${s.hata}` : s.atlandi ? "az önce çekildi, atlandı" : `${s.yeni} yeni e-posta`}</div>)}
        </div>
      )}
    </section>
  );
}
