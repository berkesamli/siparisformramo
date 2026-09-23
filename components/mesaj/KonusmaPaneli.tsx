"use client";

// Tek konuşma: başlık (kim, kanal, müşteri kartı, atama, durum), mesaj
// balonları, yanıt kutusu. "Taslak öner" yapay zekâdan metin alır ve kutuya
// yazar — gönderim yalnızca çalışanın "Gönder"iyle olur.

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import { KANAL_ADI, hesapKisa, pencereAcik, type KanalDurumu, type Konusma, type Mesaj } from "@/lib/mesaj/tur";
import type { Kullanici, Me } from "./Inbox";
import MusteriBagla from "./MusteriBagla";
import { KanalIkon, gunBasligi, zamanTam } from "./ortak";

interface MusteriOzeti {
  id: string; tur: "toptan" | "perakende"; ad: string; telefon: string; eposta: string; sehir?: string; bolge?: string;
  iskontoPct?: number; mikroBakiye?: number | null; mikroUnvan?: string; href: string;
}
interface Yanit { ok: boolean; error?: string; konusma?: Konusma; mesajlar?: Mesaj[]; musteri?: MusteriOzeti | null; pencere?: boolean; mesaj?: Mesaj; taslak?: string; yontem?: string; sablon?: string }

const tl = (n: number) => n.toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " ₺";

export default function KonusmaPaneli({ id, me, kullanicilar, taslakHazir, kanallar, hesapSira = 0, onBack, onChanged, onClosed }: {
  id: string; me: Me; kullanicilar: Kullanici[]; taslakHazir: boolean; kanallar: KanalDurumu; hesapSira?: number;
  onBack: () => void; onChanged: (k: Konusma) => void; onClosed: () => void;
}) {
  const [k, setK] = useState<Konusma | null>(null);
  const [mesajlar, setMesajlar] = useState<Mesaj[]>([]);
  const [musteri, setMusteri] = useState<MusteriOzeti | null>(null);
  const [hata, setHata] = useState("");
  const [metin, setMetin] = useState("");
  const [taslakAi, setTaslakAi] = useState(false);
  const [taslakYukleniyor, setTaslakYukleniyor] = useState(false);
  const [gonderiliyor, setGonderiliyor] = useState(false);
  const [gonderHata, setGonderHata] = useState("");
  const [gonderNotu, setGonderNotu] = useState("");
  const [baglaAcik, setBaglaAcik] = useState(false);
  const [simdi, setSimdi] = useState(() => Date.now());
  const altRef = useRef<HTMLDivElement>(null);
  const kutuRef = useRef<HTMLTextAreaElement>(null);
  const sonSayi = useRef(0);

  const yukle = useCallback(async (ilk = false) => {
    try {
      const r = await fetch(`/api/mesaj/${encodeURIComponent(id)}`, { cache: "no-store" });
      const d = (await r.json()) as Yanit;
      if (!d.ok || !d.konusma) { setHata(d.error || "Konuşma açılamadı."); return; }
      setHata("");
      setK(d.konusma);
      setMesajlar(d.mesajlar || []);
      setMusteri(d.musteri || null);
      if (ilk || (d.mesajlar || []).length !== sonSayi.current) {
        sonSayi.current = (d.mesajlar || []).length;
        setTimeout(() => altRef.current?.scrollIntoView({ block: "end", behavior: ilk ? "auto" : "smooth" }), 30);
      }
      onChanged({ ...d.konusma, okunmamis: 0 });
    } catch {
      setHata("Sunucuya ulaşılamadı.");
    }
  }, [id, onChanged]);

  useEffect(() => { void yukle(true); }, [yukle]);
  useEffect(() => {
    const t = setInterval(() => { if (document.visibilityState === "visible") { void yukle(); setSimdi(Date.now()); } }, 20_000);
    return () => clearInterval(t);
  }, [yukle]);

  // Yazı kutusu yüksekliği içerikle büyür
  useEffect(() => {
    const el = kutuRef.current; if (!el) return;
    el.style.height = "auto"; el.style.height = Math.min(220, el.scrollHeight) + "px";
  }, [metin]);

  async function patch(degisiklik: Record<string, unknown>) {
    const r = await fetch(`/api/mesaj/${encodeURIComponent(id)}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(degisiklik) });
    const d = (await r.json()) as Yanit;
    if (!d.ok || !d.konusma) { setHata(d.error || "Güncellenemedi."); return null; }
    setK(d.konusma); setMusteri(d.musteri || null); onChanged({ ...d.konusma, okunmamis: 0 });
    return d.konusma;
  }

  async function taslakAl() {
    setTaslakYukleniyor(true); setGonderHata("");
    try {
      const r = await fetch("/api/mesaj/taslak", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ konusmaId: id }) });
      const d = (await r.json()) as Yanit;
      if (!d.ok || !d.taslak) { setGonderHata(d.error || "Taslak alınamadı."); return; }
      setMetin(d.taslak); setTaslakAi(true);
      setTimeout(() => kutuRef.current?.focus(), 20);
    } catch { setGonderHata("Taslak alınamadı."); }
    finally { setTaslakYukleniyor(false); }
  }

  async function gonder() {
    const govde = metin.trim();
    if (!govde || gonderiliyor) return;
    setGonderiliyor(true); setGonderHata("");
    try {
      const r = await fetch(`/api/mesaj/${encodeURIComponent(id)}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ metin: govde, taslakAi }) });
      const d = (await r.json()) as Yanit;
      if (!d.ok || !d.mesaj) { setGonderHata(d.error || "Gönderilemedi."); return; }
      setMesajlar((m) => [...m, d.mesaj!]); sonSayi.current += 1;
      if (d.konusma) { setK(d.konusma); onChanged({ ...d.konusma, okunmamis: 0 }); }
      setMetin(""); setTaslakAi(false);
      setGonderNotu(d.yontem === "sablon" ? `Müşteri 24 saattir yazmadığı için mesaj onaylı şablonla gitti (${d.sablon}).` : "");
      setTimeout(() => altRef.current?.scrollIntoView({ block: "end", behavior: "smooth" }), 30);
    } catch { setGonderHata("Sunucuya ulaşılamadı."); }
    finally { setGonderiliyor(false); }
  }

  if (hata && !k) return <div className="ib-thread-empty"><div className="notice err">{hata}</div><button type="button" className="btn secondary small" onClick={onBack}><Icon name="chevron-left" size={15} /> Listeye dön</button></div>;
  if (!k) return <div className="ib-thread-empty"><span className="skeleton" style={{ width: 220, height: 18 }} /><span className="skeleton" style={{ width: 320, height: 14, marginTop: 8 }} /></div>;

  const pencere = pencereAcik(k, simdi);
  const kanalHazir = kanallar[k.kanal];
  const sablonla = k.kanal === "whatsapp" && !pencere && kanallar.whatsappSablon;
  const gonderilebilir = kanalHazir && (k.kanal === "email" || pencere || sablonla);
  const kimlik = k.kanal === "email" ? k.disKimlik : k.kanal === "instagram" ? (k.baslik || `Instagram ${k.disKimlik.slice(-6)}`) : `+${k.disKimlik}`;
  const iletisim = k.kanal === "whatsapp" ? `https://wa.me/${k.disKimlik}` : k.kanal === "email" ? `mailto:${k.disKimlik}` : k.baslik ? `https://instagram.com/${k.baslik.replace(/^@/, "")}` : "";

  return (
    <div className="ib-panel">
      <header className="ib-panel-head">
        <button type="button" className="btn ghost icon small ib-back" onClick={onBack} aria-label="Listeye dön"><Icon name="chevron-left" size={18} /></button>
        <span className={`ib-av ${k.kanal}`}><KanalIkon kanal={k.kanal} size={18} /></span>
        <div className="ib-panel-who">
          <strong>{k.ad || kimlik}</strong>
          <span className="muted ib-panel-sub">
            {iletisim ? <a href={iletisim} target="_blank" rel="noreferrer">{kimlik}</a> : kimlik}
            {k.kanal === "email" && k.hesap && <span className={`badge ib-hesap ib-hesap-${hesapSira}`} title={`Bu hesaba geldi, yanıt buradan gider: ${k.hesap}`}><Icon name="mail" size={10} /> {hesapKisa(k.hesap)}</span>}
          </span>
        </div>
        <div className="ib-panel-actions">
          <select className="ib-select" value={k.atanan || ""} onChange={(e) => void patch({ atanan: e.target.value })} aria-label="Atanan çalışan" title="Konuşmayı bir çalışana ata">
            <option value="">Atanmadı</option>
            {kullanicilar.map((u) => <option key={u.username} value={u.username}>{u.username === me.username ? `${u.name} (ben)` : u.name}</option>)}
          </select>
          {k.durum === "kapali" ? (
            <button type="button" className="btn secondary small" onClick={() => void patch({ durum: "acik" })}><Icon name="refresh" size={14} /> Yeniden aç</button>
          ) : (
            <button type="button" className="btn secondary small" onClick={async () => { if (await patch({ durum: "kapali" })) onClosed(); }} title="Konuşmayı kapat (müşteri tekrar yazınca açılır)"><Icon name="check-circle" size={14} /> Kapat</button>
          )}
        </div>
      </header>

      {k.kanal === "email" && k.baslik && (
        <div className="ib-konu" title="E-posta konusu"><span className="ib-konu-etiket">Konu</span><span className="ib-konu-metin">{k.baslik}</span></div>
      )}

      <div className="ib-musteri">
        {musteri ? (
          <>
            <Link href={musteri.href} className={`badge ${musteri.tur === "toptan" ? "brand" : "info"}`}><Icon name={musteri.tur === "toptan" ? "briefcase" : "user"} size={12} /> {musteri.tur === "toptan" ? "Bayi" : "Perakende"}: {musteri.ad}</Link>
            {musteri.sehir && <span className="muted">{musteri.sehir}</span>}
            {musteri.bolge && <span className="muted">· {musteri.bolge} bölgesi</span>}
            {musteri.tur === "toptan" && musteri.mikroBakiye !== undefined && (
              <span className={`badge ${musteri.mikroBakiye == null ? "" : musteri.mikroBakiye > 0 ? "warn" : "ok"}`} title="Mikro cari bakiyesi (pozitif: müşteri borçlu)">
                Mikro: {musteri.mikroBakiye == null ? "okunamadı" : tl(musteri.mikroBakiye)}
              </span>
            )}
            <button type="button" className="btn ghost xs" onClick={() => setBaglaAcik(true)}>Değiştir</button>
          </>
        ) : (
          <>
            <span className="muted"><Icon name="user" size={13} /> Müşteri defterinde kayıtlı değil.</span>
            <button type="button" className="btn secondary xs" onClick={() => setBaglaAcik(true)}><Icon name="plus" size={12} /> Müşteri bağla</button>
          </>
        )}
      </div>

      <div className={`ib-msgs ${k.kanal === "email" ? "ib-msgs-mail" : ""}`}>
        {mesajlar.map((m, i) => {
          const oncekiGun = i > 0 ? new Date(mesajlar[i - 1].at).toDateString() : "";
          const buGun = new Date(m.at).toDateString();
          return (
            <div key={m.id}>
              {oncekiGun !== buGun && <div className="ib-day"><span>{gunBasligi(m.at)}</span></div>}
              {k.kanal === "email" ? <MailKarti m={m} karsi={k.disKimlik} hesap={k.hesap} /> : <Balon m={m} />}
            </div>
          );
        })}
        {!mesajlar.length && <div className="empty">Bu konuşmada henüz mesaj yok.</div>}
        <div ref={altRef} />
      </div>

      <footer className="ib-compose">
        {!kanalHazir && <div className="notice warn">{KANAL_ADI[k.kanal]} gönderimi için kanal ayarları eksik; bu konuşmaya buradan yanıt verilemez.</div>}
        {kanalHazir && k.kanal !== "email" && !pencere && !sablonla && (
          <div className="notice warn"><Icon name="clock" size={14} /> {KANAL_ADI[k.kanal]} kuralı: müşteri son 24 saatte yazmadığı için serbest yanıt gönderilemez. Müşteri yeniden yazınca pencere açılır{k.kanal === "whatsapp" ? "; acil durumda telefon/SMS kullanın ya da Ayarlar → Mesajlar kartından onaylı şablon oluşturun" : ""}.</div>
        )}
        {sablonla && (
          <div className="notice info"><Icon name="clock" size={14} /> Müşteri son 24 saatte yazmadı: mesajınız Meta onaylı şablonun içinde gider (satır sonları tek boşluk olur, en fazla 1000 karakter); müşteri yanıtlayınca serbest yazışma açılır.</div>
        )}
        {gonderNotu && <div className="notice ok" onClick={() => setGonderNotu("")}>{gonderNotu}</div>}
        {taslakAi && <div className="ib-taslak-not"><Icon name="sparkles" size={14} /> Yapay zekâ taslağı — göndermeden önce okuyup düzeltin. <button type="button" className="btn ghost xs" onClick={() => { setMetin(""); setTaslakAi(false); }}>Temizle</button></div>}
        {gonderHata && <div className="notice err">{gonderHata}</div>}
        <div className="ib-compose-row">
          <textarea
            ref={kutuRef}
            value={metin}
            onChange={(e) => { setMetin(e.target.value); if (taslakAi && !e.target.value.trim()) setTaslakAi(false); }}
            onKeyDown={(e) => { if ((e.ctrlKey || e.metaKey) && e.key === "Enter") { e.preventDefault(); void gonder(); } }}
            placeholder={k.kanal === "email" ? `${k.hesap || "E-posta"} adresinden yanıt (konu: ${/^re:/i.test(k.baslik) ? k.baslik : "Re: " + (k.baslik || "")})` : `${KANAL_ADI[k.kanal]} yanıtı yazın… (Ctrl+Enter gönderir)`}
            rows={2}
            maxLength={sablonla ? 1000 : k.kanal === "email" ? 20000 : 4000}
            disabled={gonderiliyor}
            aria-label="Yanıt"
          />
        </div>
        <div className="ib-compose-actions">
          {taslakHazir ? (
            <button type="button" className="btn secondary small" onClick={() => void taslakAl()} disabled={taslakYukleniyor || gonderiliyor} title="Yapay zekâ, konuşmayı ve ürün/stok bilgisini okuyup yanıt taslağı önerir">
              <Icon name="sparkles" size={15} /> {taslakYukleniyor ? "Taslak yazılıyor…" : "Taslak öner"}
            </button>
          ) : (
            <span className="muted" style={{ fontSize: 12 }}>Yapay zekâ taslağı için ANTHROPIC_API_KEY gerekir.</span>
          )}
          <span className="spacer" />
          <span className="muted ib-count">{metin.length > 0 ? `${metin.length}${sablonla ? "/1000" : ""} karakter` : ""}</span>
          <button type="button" className={`btn small ${k.kanal === "whatsapp" ? "wa" : ""}`} onClick={() => void gonder()} disabled={!metin.trim() || gonderiliyor || !gonderilebilir}>
            <Icon name="arrow-up-right" size={15} /> {gonderiliyor ? "Gönderiliyor…" : k.kanal === "email" ? "E-posta gönder" : "Gönder"}
          </button>
        </div>
      </footer>

      {baglaAcik && (
        <MusteriBagla
          mevcut={musteri ? { id: musteri.id, tur: musteri.tur, ad: musteri.ad } : null}
          ipucu={k.kanal === "email" ? k.disKimlik : k.kanal === "whatsapp" ? k.disKimlik : k.ad}
          onClose={() => setBaglaAcik(false)}
          onSelect={async (id, tur) => { await patch({ musteriId: id, musteriTur: tur }); setBaglaAcik(false); }}
        />
      )}
    </div>
  );
}

function DurumIkonu({ m }: { m: Mesaj }) {
  if (m.yon !== "giden") return null;
  if (m.durum === "hata") return <span className="ib-st err" title={m.hata || "Gönderilemedi"}><Icon name="alert" size={12} /> {m.hata ? m.hata.slice(0, 80) : "hata"}</span>;
  if (m.durum === "okudu" || m.durum === "okundu") return <span className="ib-st read" title="Okundu">✓✓</span>;
  if (m.durum === "iletildi") return <span className="ib-st" title="İletildi">✓✓</span>;
  return <span className="ib-st" title="Gönderildi">✓</span>;
}

function Ekler({ ekler }: { ekler: Mesaj["ekler"] }) {
  if (!ekler?.length) return null;
  return (
    <div className="ib-ekler">
      {ekler.map((e, i) => e.tur === "image" && e.url ? (
        // eslint-disable-next-line @next/next/no-img-element
        <a key={i} href={e.url} target="_blank" rel="noreferrer" className="ib-ek-img"><img src={e.url} alt={e.ad} loading="lazy" /></a>
      ) : (
        <a key={i} href={e.url || "#"} target={e.url ? "_blank" : undefined} rel="noreferrer" className={`ib-ek ${e.url ? "" : "off"}`} title={e.url ? "Aç" : "Ek kaydedilemedi (depo yok)"}>
          <Icon name={e.tur === "audio" ? "activity" : e.tur === "video" || e.tur === "image" ? "image" : "file-text"} size={14} /> {e.ad}{e.boyut ? <small> · {Math.round(e.boyut / 1024)} KB</small> : null}
        </a>
      ))}
    </div>
  );
}

/** WhatsApp / Instagram: sohbet balonu. */
function Balon({ m }: { m: Mesaj }) {
  return (
    <div className={`ib-bubble ${m.yon}`}>
      {m.yon === "giden" && <span className="ib-bubble-who">{m.gonderen || "Biz"}{m.taslakAi && <span className="ib-ai" title="Yapay zekâ taslağından gönderildi"><Icon name="sparkles" size={10} /> taslak</span>}</span>}
      {m.govde && <div className="ib-bubble-text">{m.govde}</div>}
      <Ekler ekler={m.ekler} />
      <span className="ib-bubble-meta">{zamanTam(m.at)} <DurumIkonu m={m} /></span>
    </div>
  );
}

const MAIL_KIRP_KARAKTER = 900;
const MAIL_KIRP_SATIR = 14;

/** Gövde "Konu: …" ile başlıyorsa (konu değişmiş e-posta) konuyu ayırır; fazla boş satırları toplar. */
function mailGovde(govde: string): { konu: string; metin: string } {
  let konu = "";
  let metin = govde || "";
  const m = metin.match(/^Konu: (.+)\n\n?/);
  if (m) { konu = m[1].trim(); metin = metin.slice(m[0].length); }
  metin = metin.replace(/\r\n?/g, "\n").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  return { konu, metin };
}

const basHarfler = (ad: string) => ad.split(/\s+/).filter(Boolean).slice(0, 2).map((x) => x[0]?.toLocaleUpperCase("tr-TR") || "").join("") || "?";

/** E-posta: sohbet balonu yerine okunaklı mektup kartı (gönderen, saat, konu, gövde, ekler). */
function MailKarti({ m, karsi, hesap }: { m: Mesaj; karsi: string; hesap: string }) {
  const [acik, setAcik] = useState(false);
  const { konu, metin } = mailGovde(m.govde);
  const uzun = metin.length > MAIL_KIRP_KARAKTER || metin.split("\n").length > MAIL_KIRP_SATIR;
  const giden = m.yon === "giden";
  const ad = m.gonderen || (giden ? "Biz" : karsi);
  return (
    <article className={`ib-mail ${m.yon}`}>
      <header className="ib-mail-head">
        <span className={`ib-mail-av ${m.yon}`} aria-hidden>{giden ? <Icon name="arrow-up-right" size={14} /> : basHarfler(ad)}</span>
        <div className="ib-mail-who">
          <strong>{ad}{giden && m.taslakAi && <span className="ib-ai" title="Yapay zekâ taslağından gönderildi"><Icon name="sparkles" size={10} /> taslak</span>}</strong>
          <span className="muted">{giden ? <>{hesap || "biz"} → {karsi}</> : karsi}</span>
        </div>
        <span className="ib-mail-time">{zamanTam(m.at)} <DurumIkonu m={m} /></span>
      </header>
      {konu && <div className="ib-mail-konu"><span>Konu</span>{konu}</div>}
      <div className={`ib-mail-body ${uzun && !acik ? "kirpik" : ""}`}>{metin || <em className="muted">(metin yok)</em>}</div>
      {uzun && (
        <button type="button" className={`btn ghost xs ib-mail-more ${acik ? "acik" : ""}`} onClick={() => setAcik(!acik)}>
          <Icon name="chevron-down" size={13} /> {acik ? "Daralt" : "Devamını göster"}
        </button>
      )}
      <Ekler ekler={m.ekler} />
    </article>
  );
}
