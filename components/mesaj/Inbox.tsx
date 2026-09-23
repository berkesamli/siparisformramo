"use client";

// Gelen kutusu: sol liste (filtre + arama), sağ konuşma paneli. Telefonda tek
// sütun: liste → konuşma (geri tuşu). Liste 30 sn'de bir, Gmail 60 sn'de bir tazelenir.

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Icon from "@/components/shell/Icon";
import { KANAL_ADI, hesapKisa, type Kanal, type KanalDurumu, type Konusma, type KonusmaDurum } from "@/lib/mesaj/tur";
import KonusmaPaneli from "./KonusmaPaneli";
import YeniMesaj from "./YeniMesaj";
import { KanalIkon, zamanKisa } from "./ortak";

export interface Me { username: string; name: string; owner: boolean }
export interface Kullanici { username: string; name: string }

interface ListeYaniti {
  ok: boolean; error?: string; kurulum?: boolean;
  konusmalar?: Konusma[]; okunmamis?: number; kanallar?: Record<Kanal, boolean>; taslak?: boolean; kullanicilar?: Kullanici[];
  hesaplar?: { email: string[] };
  senk?: { hesap: string; yeni: number; hata?: string; atlandi?: boolean }[];
}

const DURUMLAR: { v: KonusmaDurum | ""; ad: string }[] = [
  { v: "", ad: "Tümü" }, { v: "acik", ad: "Açık" }, { v: "yanitlandi", ad: "Yanıtlandı" }, { v: "kapali", ad: "Kapalı" },
];
const KANALLAR: Kanal[] = ["whatsapp", "instagram", "email"];

export default function Inbox({ me, dbHazir, kanallar, taslak, ilkKonusma }: {
  me: Me; dbHazir: boolean; kanallar: KanalDurumu; taslak: boolean; ilkKonusma: string;
}) {
  const [yeniAcik, setYeniAcik] = useState(false);
  const [bilgi, setBilgi] = useState("");
  const [kanal, setKanal] = useState<Kanal | "">("");
  const [hesap, setHesap] = useState("");
  const [epostaHesaplar, setEpostaHesaplar] = useState<string[]>([]);
  const [durum, setDurum] = useState<KonusmaDurum | "">("");
  const [bana, setBana] = useState(false);
  const [q, setQ] = useState("");
  const [liste, setListe] = useState<Konusma[]>([]);
  const [yukleniyor, setYukleniyor] = useState(true);
  const [hata, setHata] = useState("");
  const [kullanicilar, setKullanicilar] = useState<Kullanici[]>([]);
  const [senkNotu, setSenkNotu] = useState("");
  const [secili, setSecili] = useState<string>(ilkKonusma);
  const sonSenk = useRef(0);
  const qRef = useRef(q);
  qRef.current = q;
  const ilkArama = useRef(true);

  const yukle = useCallback(async (sessiz = false, zorla = false) => {
    if (!sessiz) setYukleniyor(true);
    const p = new URLSearchParams();
    if (kanal) p.set("kanal", kanal);
    if (kanal === "email" && hesap) p.set("hesap", hesap);
    if (durum) p.set("durum", durum);
    if (bana) p.set("atanan", "ben");
    if (qRef.current.trim()) p.set("q", qRef.current.trim());
    const senk = zorla || Date.now() - sonSenk.current > 60_000;
    if (senk) { p.set("senk", "1"); if (zorla) p.set("zorla", "1"); sonSenk.current = Date.now(); }
    try {
      const r = await fetch(`/api/mesaj/konusmalar?${p}`, { cache: "no-store" });
      const d = (await r.json()) as ListeYaniti;
      if (!d.ok) { setHata(d.error || "Liste alınamadı."); return; }
      setHata("");
      setListe(d.konusmalar || []);
      if (d.kullanicilar) setKullanicilar(d.kullanicilar);
      if (d.hesaplar?.email) setEpostaHesaplar(d.hesaplar.email);
      if (d.senk?.length) {
        const hatali = d.senk.filter((s) => s.hata);
        const yeni = d.senk.reduce((n, s) => n + (s.yeni || 0), 0);
        setSenkNotu(hatali.length ? `Okunamadı: ${hatali.map((s) => `${s.hesap} — ${s.hata}`).join("; ")}` : yeni ? `${yeni} yeni mesaj alındı (${d.senk.filter((s) => s.yeni).map((s) => `${s.hesap === "Instagram" ? "Instagram" : hesapKisa(s.hesap)}: ${s.yeni}`).join(", ")}).` : "");
      }
    } catch {
      setHata("Sunucuya ulaşılamadı.");
    } finally {
      setYukleniyor(false);
    }
  }, [kanal, hesap, durum, bana]);

  useEffect(() => { if (dbHazir) void yukle(); }, [yukle, dbHazir]);
  // Arama: yazmayı bırakınca (ilk render'da yükleme zaten yukarıda yapılır)
  useEffect(() => {
    if (!dbHazir) return;
    if (ilkArama.current) { ilkArama.current = false; return; }
    const t = setTimeout(() => void yukle(true), 350);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q]);
  // Arka planda tazele (sekme görünürken)
  useEffect(() => {
    if (!dbHazir) return;
    const t = setInterval(() => { if (document.visibilityState === "visible") void yukle(true); }, 30_000);
    return () => clearInterval(t);
  }, [yukle, dbHazir]);

  const konusmaGuncellendi = useCallback((k: Konusma) => {
    setListe((l) => {
      const i = l.findIndex((x) => x.id === k.id);
      if (i < 0) return l;
      const n = [...l]; n[i] = { ...n[i], ...k }; return n;
    });
  }, []);

  const toplamOkunmamis = useMemo(() => liste.reduce((n, k) => n + (k.okunmamis || 0), 0), [liste]);
  const hicKanalYok = !kanallar.whatsapp && !kanallar.instagram && !kanallar.email;

  if (!dbHazir) return <KurulumKarti me={me} kanallar={kanallar} />;

  return (
    <div className={`ib ${secili ? "thread-open" : ""}`}>
      <aside className="ib-list">
        <div className="ib-list-head">
          <div className="ib-title">
            <span className="page-head-icon" aria-hidden><Icon name="inbox" size={20} /></span>
            <div>
              <h1>Mesajlar</h1>
              <span className="muted">{toplamOkunmamis > 0 ? `${toplamOkunmamis} okunmamış` : "Gelen kutusu"}</span>
            </div>
            {kanallar.whatsapp && kanallar.whatsappSablon && (
              <button type="button" className="btn wa small ib-yeni-btn" title="Bizim başlattığımız WhatsApp mesajı (onaylı şablonla)" onClick={() => setYeniAcik(true)}>
                <Icon name="edit" size={15} /> <span>Yeni</span>
              </button>
            )}
            <button type="button" className={`btn ghost icon small ${yukleniyor ? "spin" : ""}`} title="Yenile" aria-label="Yenile" onClick={() => void yukle(false, true)}>
              <Icon name="refresh" size={17} />
            </button>
          </div>
          <div className="ib-search">
            <Icon name="search" size={16} />
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Ad, numara, e-posta, konu…" aria-label="Konuşma ara" />
            {q && <button type="button" className="ib-search-clear" aria-label="Temizle" onClick={() => setQ("")}><Icon name="x" size={14} /></button>}
          </div>
          <div className="ib-chips">
            <button type="button" className={`chip ${kanal === "" ? "sel" : ""}`} onClick={() => setKanal("")}>Tümü</button>
            {KANALLAR.map((k) => (
              <button key={k} type="button" className={`chip ib-chip-${k} ${kanal === k ? "sel" : ""}`} onClick={() => setKanal(kanal === k ? "" : k)} title={kanallar[k] ? undefined : "Bu kanal henüz bağlı değil"}>
                <KanalIkon kanal={k} size={13} /> {KANAL_ADI[k]}{!kanallar[k] && <span className="ib-chip-off" aria-label="bağlı değil" />}
              </button>
            ))}
          </div>
          {kanal === "email" && epostaHesaplar.length > 1 && (
            <div className="ib-chips ib-hesaplar" aria-label="E-posta hesabı">
              <button type="button" className={`chip ${hesap === "" ? "sel" : ""}`} onClick={() => setHesap("")}>Tüm hesaplar</button>
              {epostaHesaplar.map((h, i) => (
                <button key={h} type="button" className={`chip ib-hesap-${i % 3} ${hesap === h ? "sel" : ""}`} onClick={() => setHesap(hesap === h ? "" : h)} title={h}>
                  <Icon name="mail" size={12} /> {hesapKisa(h)}
                </button>
              ))}
            </div>
          )}
          <div className="ib-filters">
            <div className="seg ib-seg">
              {DURUMLAR.map((d) => <button key={d.v} type="button" className={durum === d.v ? "active" : ""} onClick={() => setDurum(d.v)}>{d.ad}</button>)}
            </div>
            <button type="button" className={`chip ${bana ? "sel" : ""}`} onClick={() => setBana(!bana)}><Icon name="user" size={13} /> Bana atanan</button>
          </div>
        </div>

        {hata && <div className="notice err" style={{ margin: "8px 12px" }}>{hata}</div>}
        {senkNotu && <div className={`notice ${/okunamadı/.test(senkNotu) ? "warn" : "ok"} ib-senk`} onClick={() => setSenkNotu("")}>{senkNotu}</div>}
        {bilgi && <div className="notice ok ib-senk" onClick={() => setBilgi("")}>{bilgi}</div>}
        {hicKanalYok && !hata && (
          <div className="notice warn" style={{ margin: "8px 12px" }}>Henüz hiçbir kanal bağlı değil. Bildirim Ayarları / Vercel ortam değişkenleriyle WhatsApp, Instagram ve Gmail bağlanınca mesajlar burada toplanır.</div>
        )}

        <div className={`ib-rows ${yukleniyor && liste.length ? "loading-dim" : ""}`} role="list">
          {yukleniyor && !liste.length && [0, 1, 2, 3, 4].map((i) => <div key={i} className="ib-row skeleton-row"><span className="skeleton" style={{ width: 40, height: 40, borderRadius: 12 }} /><span><span className="skeleton" style={{ width: "55%" }} /><span className="skeleton" style={{ width: "85%", marginTop: 6 }} /></span></div>)}
          {!yukleniyor && !liste.length && (
            <div className="empty">
              <div className="empty-icon"><Icon name="inbox" size={24} /></div>
              <strong>{q || kanal || durum || bana ? "Bu süzgeçle konuşma yok" : "Henüz mesaj yok"}</strong>
              {q || kanal || durum || bana ? "Süzgeçleri temizleyip tekrar bakın." : "Müşteriler yazdıkça konuşmalar burada listelenir."}
            </div>
          )}
          {liste.map((k) => (
            <KonusmaSatiri key={k.id} k={k} secili={k.id === secili} kullanicilar={kullanicilar} hesapSira={k.kanal === "email" ? Math.max(0, epostaHesaplar.indexOf(k.hesap)) % 3 : 0} onClick={() => setSecili(k.id)} />
          ))}
        </div>
      </aside>

      {yeniAcik && (
        <YeniMesaj
          kanallar={kanallar}
          onClose={() => setYeniAcik(false)}
          onSent={(k, not) => {
            setYeniAcik(false);
            setListe((l) => [k, ...l.filter((x) => x.id !== k.id)]);
            setSecili(k.id);
            setBilgi(not);
          }}
        />
      )}

      <section className="ib-thread">
        {secili ? (
          <KonusmaPaneli
            key={secili}
            id={secili}
            me={me}
            kullanicilar={kullanicilar}
            taslakHazir={taslak}
            kanallar={kanallar}
            hesapSira={Math.max(0, epostaHesaplar.indexOf(liste.find((x) => x.id === secili)?.hesap || "")) % 3}
            onBack={() => setSecili("")}
            onChanged={konusmaGuncellendi}
            onClosed={() => { setSecili(""); void yukle(true); }}
          />
        ) : (
          <div className="ib-thread-empty">
            <div className="empty-icon"><Icon name="message" size={26} /></div>
            <strong>Bir konuşma seçin</strong>
            <span className="muted">Soldaki listeden bir müşteri seçince yazışma burada açılır. Yapay zekâ yalnızca taslak önerir; göndermek sizin elinizde.</span>
          </div>
        )}
      </section>
    </div>
  );
}

function KonusmaSatiri({ k, secili, kullanicilar, hesapSira, onClick }: { k: Konusma; secili: boolean; kullanicilar: Kullanici[]; hesapSira: number; onClick: () => void }) {
  const atananAd = k.atanan ? (kullanicilar.find((u) => u.username === k.atanan)?.name || k.atanan) : "";
  const altBilgi = k.kanal === "email" ? (k.baslik || "(konu yok)") : k.kanal === "instagram" ? (k.baslik || "Instagram") : `+${k.disKimlik}`;
  const ozet = (k.sonMesajOzet || "").replace(/^Konu: .*\n+/, "").replace(/\s+/g, " ").trim() || "—";
  return (
    <button type="button" role="listitem" className={`ib-row ${secili ? "sel" : ""} ${k.okunmamis ? "unread" : ""} ${k.durum}`} onClick={onClick}>
      <span className={`ib-av ${k.kanal}`}><KanalIkon kanal={k.kanal} size={17} /></span>
      <span className="ib-row-main">
        <span className="ib-row-top">
          <span className="ib-row-name">{k.ad || k.disKimlik}</span>
          {k.musteriTur && <span className={`ib-tag ${k.musteriTur}`}>{k.musteriTur === "toptan" ? "Bayi" : "Perakende"}</span>}
          <span className="ib-row-time">{zamanKisa(k.sonMesajAt)}</span>
        </span>
        <span className="ib-row-sub">
          {k.kanal === "email" && k.hesap && <span className={`ib-hesap-nokta ib-hesap-${hesapSira}`} title={`Hesap: ${k.hesap}`}>{hesapKisa(k.hesap, 14)}</span>}
          <span className="ib-row-sub-text">{altBilgi}</span>
        </span>
        <span className="ib-row-ozet">
          {k.durum === "yanitlandi" && <Icon name="check-circle" size={12} />}
          {k.durum === "kapali" && <span className="ib-tag kapali">Kapalı</span>}
          <span className="ib-row-ozet-text">{ozet}</span>
        </span>
      </span>
      <span className="ib-row-side">
        {k.okunmamis > 0 && <span className="badge count ib-unread">{k.okunmamis > 99 ? "99+" : k.okunmamis}</span>}
        {atananAd && <span className="ib-atanan" title={`Atanan: ${atananAd}`}>{atananAd.split(/\s+/).slice(0, 2).map((x) => x[0]?.toLocaleUpperCase("tr-TR")).join("")}</span>}
      </span>
    </button>
  );
}

function KurulumKarti({ me, kanallar }: { me: Me; kanallar: Record<Kanal, boolean> }) {
  return (
    <div className="card ib-setup">
      <div className="card-head">
        <span className="card-head-icon"><Icon name="inbox" size={20} /></span>
        <div><h2>Mesajlar henüz kurulmadı</h2><div className="card-head-sub">WhatsApp, Instagram ve e-posta mesajlarını tek ekranda toplamak için veri tabanı bağlantısı gerekir.</div></div>
      </div>
      {me.owner ? (
        <ol className="ib-setup-steps">
          <li><strong>Veri tabanı:</strong> Vercel → Storage → <em>Neon Postgres</em> oluşturup projeye bağlayın; <code>DATABASE_URL</code> kendiliğinden eklenir. Tablolar ilk açılışta kurulur.</li>
          <li><strong>WhatsApp:</strong> mevcut Cloud API bağlantısı yeterlidir (webhook <code>/api/whatsapp/webhook</code>, alan: <em>messages</em>). {kanallar.whatsapp ? "✓ Bağlı." : "Henüz bağlı değil."}</li>
          <li><strong>Gmail:</strong> her hesap için Google &quot;uygulama şifresi&quot; alın ve <code>GMAIL_HESAPLAR=&quot;adres:şifre;adres2:şifre2&quot;</code> girin. {kanallar.email ? "✓ Bağlı." : "Henüz bağlı değil."}</li>
          <li><strong>Instagram:</strong> Meta uygulamasına <em>Instagram</em> ürünü + <em>instagram_manage_messages</em> izni (App Review), webhook <code>/api/mesaj/webhook</code>; <code>INSTAGRAM_TOKEN</code>, <code>INSTAGRAM_PAGE_ID</code>, <code>INSTAGRAM_ACCOUNT_ID</code>. {kanallar.instagram ? "✓ Bağlı." : "Henüz bağlı değil."}</li>
          <li><strong>Kimler görsün:</strong> <code>MESAJ_USERNAMES=berke,ramazan</code> (boşsa yalnızca sahipler).</li>
        </ol>
      ) : (
        <p className="muted">Kurulum için firma sahibine başvurun.</p>
      )}
    </div>
  );
}
