"use client";

// Bizim başlattığımız WhatsApp mesajı: müşteri defterinden seç ya da numara yaz, metni gönder.
// Müşteri son 24 saatte yazmadıysa mesaj Meta onaylı şablonla gider (sunucu karar verir).

import { useEffect, useMemo, useState } from "react";
import Icon from "@/components/shell/Icon";
import { eslesir } from "@/lib/search-norm";
import type { KanalDurumu, Konusma } from "@/lib/mesaj/tur";
import { KanalIkon } from "./ortak";

interface Toptan { id: string; firstName: string; lastName: string; company: string; phone: string; city: string }
interface Perakende { id: string; name: string; phone: string }
interface Alici { telefon: string; ad: string; musteriId: string | null; musteriTur: "toptan" | "perakende" | null }

const telefonGecerli = (t: string) => String(t || "").replace(/\D/g, "").length >= 10;

export default function YeniMesaj({ kanallar, onClose, onSent }: {
  kanallar: KanalDurumu; onClose: () => void; onSent: (k: Konusma, not: string) => void;
}) {
  const [q, setQ] = useState("");
  const [toptan, setToptan] = useState<Toptan[]>([]);
  const [perakende, setPerakende] = useState<Perakende[]>([]);
  const [yukleniyor, setYukleniyor] = useState(true);
  const [alici, setAlici] = useState<Alici | null>(null);
  const [metin, setMetin] = useState("");
  const [gonderiliyor, setGonderiliyor] = useState(false);
  const [hata, setHata] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const [a, b] = await Promise.all([
          fetch("/api/musteriler", { cache: "no-store" }).then((r) => r.json()).catch(() => ({})),
          fetch("/api/perakende/musteriler", { cache: "no-store" }).then((r) => r.json()).catch(() => ({})),
        ]);
        setToptan((Array.isArray(a?.customers) ? a.customers : []).filter((c: Toptan) => telefonGecerli(c.phone)));
        setPerakende((Array.isArray(b?.customers) ? b.customers : []).filter((c: Perakende) => telefonGecerli(c.phone)));
      } finally { setYukleniyor(false); }
    })();
  }, []);

  const s = q.trim();
  const t = useMemo(() => (s ? toptan.filter((c) => eslesir(s, c.company, c.firstName, c.lastName, c.phone, c.city)) : toptan).slice(0, 20), [s, toptan]);
  const p = useMemo(() => (s ? perakende.filter((c) => eslesir(s, c.name, c.phone)) : perakende).slice(0, 20), [s, perakende]);
  const numaraGibi = telefonGecerli(s) && /^[\d\s()+-]+$/.test(s);

  async function gonder() {
    if (!alici || !metin.trim() || gonderiliyor) return;
    setGonderiliyor(true); setHata("");
    try {
      const r = await fetch("/api/mesaj/yeni", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ telefon: alici.telefon, ad: alici.ad, metin: metin.trim(), musteriId: alici.musteriId, musteriTur: alici.musteriTur }) });
      const d = (await r.json()) as { ok: boolean; error?: string; konusma?: Konusma; yontem?: string; sablon?: string };
      if (!d.ok || !d.konusma) { setHata(d.error || "Gönderilemedi."); return; }
      onSent(d.konusma, d.yontem === "sablon" ? `Şablonla gönderildi (${d.sablon}); müşteri yanıt verince serbest yazışma açılır.` : "Gönderildi.");
    } catch { setHata("Sunucuya ulaşılamadı."); }
    finally { setGonderiliyor(false); }
  }

  return (
    <div className="backdrop ib-modal-wrap" onClick={onClose} role="dialog" aria-modal="true" aria-label="Yeni WhatsApp mesajı">
      <div className="ib-modal" onClick={(e) => e.stopPropagation()}>
        <div className="ib-modal-head">
          <strong><span className="ib-av whatsapp ib-av-sm"><KanalIkon kanal="whatsapp" size={14} /></span> Yeni WhatsApp mesajı</strong>
          <button type="button" className="btn ghost icon small" onClick={onClose} aria-label="Kapat"><Icon name="x" size={16} /></button>
        </div>

        {!kanallar.whatsapp && <div className="notice err" style={{ margin: "0 0 10px" }}>WhatsApp Cloud API ayarlı değil; gönderim yapılamaz.</div>}
        {kanallar.whatsapp && !kanallar.whatsappSablon && (
          <div className="notice warn" style={{ margin: "0 0 10px" }}>Onaylı şablon tanımlı değil (<code>WHATSAPP_TEMPLATE_SERBEST</code>): yalnızca son 24 saatte bize yazmış müşterilere gönderilebilir.</div>
        )}

        {!alici ? (
          <>
            <div className="ib-search" style={{ margin: "0 0 10px" }}>
              <Icon name="search" size={16} />
              <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Müşteri adı, firma ya da numara yazın…" aria-label="Alıcı ara" />
            </div>
            <div className={`ib-modal-body ${gonderiliyor ? "loading-dim" : ""}`}>
              {numaraGibi && (
                <button type="button" className="ib-modal-row" onClick={() => setAlici({ telefon: s, ad: "", musteriId: null, musteriTur: null })}>
                  <span className="badge"><Icon name="phone" size={11} /></span>
                  <span><strong>{s}</strong><small>Defterde olmayan numaraya gönder</small></span>
                </button>
              )}
              {yukleniyor && <div className="empty">Müşteriler yükleniyor…</div>}
              {!yukleniyor && !t.length && !p.length && !numaraGibi && <div className="empty">Eşleşen müşteri yok. Numarayı doğrudan yazabilirsiniz (örn. 0532 111 22 33).</div>}
              {t.length > 0 && <div className="ib-modal-grp">Bayiler (toptan)</div>}
              {t.map((c) => (
                <button key={c.id} type="button" className="ib-modal-row" onClick={() => setAlici({ telefon: c.phone, ad: c.company || `${c.firstName} ${c.lastName}`.trim(), musteriId: c.id, musteriTur: "toptan" })}>
                  <span className="badge brand"><Icon name="briefcase" size={11} /></span>
                  <span><strong>{c.company || `${c.firstName} ${c.lastName}`.trim()}</strong><small>{[c.city, c.phone].filter(Boolean).join(" · ")}</small></span>
                </button>
              ))}
              {p.length > 0 && <div className="ib-modal-grp">Perakende müşteriler</div>}
              {p.map((c) => (
                <button key={c.id} type="button" className="ib-modal-row" onClick={() => setAlici({ telefon: c.phone, ad: c.name, musteriId: c.id, musteriTur: "perakende" })}>
                  <span className="badge info"><Icon name="user" size={11} /></span>
                  <span><strong>{c.name}</strong><small>{c.phone}</small></span>
                </button>
              ))}
            </div>
          </>
        ) : (
          <>
            <div className="ib-alici">
              <span className={`badge ${alici.musteriTur === "toptan" ? "brand" : alici.musteriTur === "perakende" ? "info" : ""}`}>
                <Icon name={alici.musteriTur === "toptan" ? "briefcase" : "user"} size={11} /> {alici.ad || "Kayıtsız numara"}
              </span>
              <span className="muted">{alici.telefon}</span>
              <span className="spacer" />
              <button type="button" className="btn ghost xs" onClick={() => setAlici(null)}>Değiştir</button>
            </div>
            <textarea
              autoFocus
              className="ib-yeni-metin"
              rows={5}
              value={metin}
              onChange={(e) => setMetin(e.target.value)}
              onKeyDown={(e) => { if ((e.ctrlKey || e.metaKey) && e.key === "Enter") { e.preventDefault(); void gonder(); } }}
              placeholder="Mesajınız… (Ctrl+Enter gönderir)"
              maxLength={1000}
              aria-label="Mesaj"
            />
            <p className="muted" style={{ fontSize: 12, margin: "6px 0 10px" }}>
              Müşteri son 24 saatte bize yazdıysa mesaj olduğu gibi gider; yazmadıysa Meta onaylı şablonun içinde gider ve müşteri yanıtlayınca serbest yazışma açılır.
            </p>
            {hata && <div className="notice err" style={{ margin: "0 0 10px" }}>{hata}</div>}
            <div className="ib-compose-actions">
              <span className="muted ib-count">{metin.length}/1000</span>
              <span className="spacer" />
              <button type="button" className="btn secondary small" onClick={onClose} disabled={gonderiliyor}>Vazgeç</button>
              <button type="button" className="btn wa small" onClick={() => void gonder()} disabled={!metin.trim() || gonderiliyor || !kanallar.whatsapp}>
                <Icon name="arrow-up-right" size={15} /> {gonderiliyor ? "Gönderiliyor…" : "Gönder"}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
