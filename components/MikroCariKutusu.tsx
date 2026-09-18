"use client";

// Müşterinin Mikro Jump'taki (resmi) cari kartı: eşleştirme ve canlı bakiye.
// - Cari kartta tam kart (arama, bağla/kaldır, bakiye kutuları)
// - Sipariş formunda compact: müşteri seçilince tek satır bakiye uyarısı
// Mikro bağlantısı ayarlı değilse hiçbir şey çizmez.

import { useCallback, useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Ozet {
  cariKod: string; unvan: string; borc: number; alacak: number; bakiye: number;
  vadesiGecen: number; sonHareket: string | null; vadeVar?: boolean;
}
interface Durum {
  ok: boolean; kurulu?: boolean; bagli?: boolean; cariKod?: string; unvan?: string;
  ozet?: Ozet; hata?: string; error?: string; oneri?: string;
}
interface Kart { cariKod: string; unvan: string; unvan2?: string }

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const bakiyeAciklama = (b: number) => (b > 0 ? "müşteri borçlu" : b < 0 ? "müşteri alacaklı" : "bakiye yok");

export default function MikroCariKutusu({
  customerId,
  compact = false,
  onChange,
}: {
  customerId: string;
  compact?: boolean;
  onChange?: () => void;
}) {
  const [durum, setDurum] = useState<Durum | null>(null);
  const [yukleniyor, setYukleniyor] = useState(true);
  const [aramaAcik, setAramaAcik] = useState(false);
  const [q, setQ] = useState("");
  const [sonuclar, setSonuclar] = useState<Kart[] | null>(null);
  const [ariyor, setAriyor] = useState(false);
  const [aramaHata, setAramaHata] = useState("");
  const [kaydediyor, setKaydediyor] = useState(false);

  const ara = useCallback(async (metin: string) => {
    const m = metin.trim();
    if (m.length < 2) { setAramaHata("En az iki harf yazın."); return; }
    setAriyor(true); setAramaHata(""); setSonuclar(null);
    try {
      const r = await fetch(`/api/mikro/cari?q=${encodeURIComponent(m)}`);
      const d = await r.json();
      if (d.ok) setSonuclar(d.cariler || []);
      else setAramaHata(d.hata || d.error || "Arama başarısız.");
    } catch { setAramaHata("Sunucuya ulaşılamadı."); }
    finally { setAriyor(false); }
  }, []);

  const yukle = useCallback(async () => {
    setYukleniyor(true);
    try {
      const r = await fetch(`/api/mikro/cari?musteri=${encodeURIComponent(customerId)}`);
      const d = (await r.json()) as Durum;
      setDurum(d);
      if (d.oneri) setQ((eski) => eski || d.oneri || "");
      // Tam kartta bağlı değilse önerilen adla hemen aranır; çalışan tek tıkla bağlar
      if (!compact && d.ok && d.bagli === false && d.oneri) { setAramaAcik(true); ara(d.oneri); }
    } catch {
      setDurum({ ok: false, hata: "Sunucuya ulaşılamadı." });
    } finally {
      setYukleniyor(false);
    }
  }, [customerId, compact, ara]);

  useEffect(() => {
    setDurum(null); setSonuclar(null); setAramaAcik(false); setQ("");
    yukle();
  }, [yukle]);

  async function esle(kart: Kart | null) {
    setKaydediyor(true);
    try {
      const r = await fetch("/api/mikro/cari", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ musteri: customerId, cariKod: kart?.cariKod || "", unvan: kart?.unvan || "" }),
      });
      const d = await r.json();
      if (!d.ok) { setAramaHata(d.error || "Kaydedilemedi."); return; }
      setAramaAcik(false); setSonuclar(null);
      await yukle();
      onChange?.();
    } catch { setAramaHata("Sunucuya ulaşılamadı."); }
    finally { setKaydediyor(false); }
  }

  // Mikro kurulu değilse (ya da yetki yoksa) sessizce gizlenir
  if (durum && durum.kurulu === false) return null;
  if (durum && !durum.ok && durum.error === "Yetkisiz") return null;

  const ozet = durum?.ozet;

  const aramaKutusu = (
    <div style={{ marginTop: 8 }}>
      <div className="row" style={{ gap: 8, flexWrap: "wrap" }}>
        <input
          style={{ flex: "1 1 220px", minWidth: 0 }}
          placeholder="Mikro'daki firma adı ya da cari kodu"
          value={q}
          onChange={(e) => setQ(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); ara(q); } }}
          autoComplete="off"
        />
        <button type="button" className="btn secondary" onClick={() => ara(q)} disabled={ariyor}>
          <Icon name="search" size={15} /> {ariyor ? "Aranıyor…" : "Mikro'da ara"}
        </button>
        {durum?.bagli && (
          <button type="button" className="btn secondary" onClick={() => setAramaAcik(false)}>Vazgeç</button>
        )}
      </div>
      {aramaHata && <div className="notice err" style={{ marginTop: 8 }}>{aramaHata}</div>}
      {sonuclar && sonuclar.length === 0 && (
        <div className="muted" style={{ fontSize: 13, marginTop: 8 }}>
          Mikro&apos;da eşleşen cari bulunamadı. Firma adının bir bölümünü (tek kelime) yazıp tekrar deneyin.
        </div>
      )}
      {sonuclar && sonuclar.length > 0 && (
        <ul style={{ listStyle: "none", margin: "8px 0 0", padding: 0 }}>
          {sonuclar.map((k) => (
            <li key={k.cariKod} className="row" style={{ gap: 10, padding: "7px 0", borderTop: "1px solid var(--border)", alignItems: "center" }}>
              <span style={{ minWidth: 0, flex: "1 1 200px" }}>
                <strong style={{ fontSize: 13.5 }}>{k.unvan || "(ünvansız)"}</strong>
                {k.unvan2 && <span className="muted" style={{ fontSize: 12.5 }}> · {k.unvan2}</span>}
                <span className="muted" style={{ display: "block", fontSize: 12 }}><code>{k.cariKod}</code></span>
              </span>
              <button type="button" className="btn small" onClick={() => esle(k)} disabled={kaydediyor}>
                <Icon name="check-circle" size={14} /> Bu cariyi bağla
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );

  // ---------- Sipariş formu: tek satır ----------
  if (compact) {
    if (yukleniyor && !durum) return <div className="muted" style={{ fontSize: 12.5, marginTop: 6 }}>Mikro bakiyesi alınıyor…</div>;
    if (!durum) return null;
    if (durum.bagli && ozet) {
      const sinif = ozet.vadesiGecen > 0 ? "err" : ozet.bakiye > 0 ? "warn" : "ok";
      return (
        <div className={`notice ${sinif}`} style={{ marginTop: 8, fontSize: 13, display: "flex", flexWrap: "wrap", gap: "4px 12px", alignItems: "center" }}>
          <span><strong>Mikro bakiye:</strong> ₺{fmt(ozet.bakiye)} <span className="muted">({bakiyeAciklama(ozet.bakiye)})</span></span>
          {ozet.vadeVar !== false && ozet.vadesiGecen > 0 && <span><strong>Vadesi geçen:</strong> ₺{fmt(ozet.vadesiGecen)}</span>}
          {ozet.sonHareket && <span className="muted">Son hareket {ozet.sonHareket.split("-").reverse().join(".")}</span>}
          <span className="muted" style={{ fontSize: 12 }}>{durum.unvan || ozet.unvan} · {durum.cariKod}</span>
        </div>
      );
    }
    if (durum.bagli && !ozet) {
      return <div className="notice warn" style={{ marginTop: 8, fontSize: 13 }}>Mikro bakiyesi okunamadı: {durum.hata || "bilinmeyen hata"}</div>;
    }
    return (
      <div style={{ marginTop: 8 }}>
        {!aramaAcik ? (
          <button type="button" className="btn small secondary" onClick={() => { setAramaAcik(true); if (q) ara(q); }}>
            <Icon name="briefcase" size={14} /> Mikro cari kartıyla eşleştir (bakiye görünsün)
          </button>
        ) : aramaKutusu}
      </div>
    );
  }

  // ---------- Cari kart: tam kart ----------
  return (
    <div className="card" style={{ marginTop: 18 }}>
      <div className="card-head">
        <span className="card-head-icon"><Icon name="briefcase" size={16} /></span>
        <div>
          <h2>Mikro (Resmi) Cari</h2>
          <span className="card-head-sub">
            {durum?.bagli ? <>{durum.unvan || ozet?.unvan} · <code>{durum.cariKod}</code></> : "Ankara sunucusundaki Mikro Jump kaydı · canlı bakiye"}
          </span>
        </div>
        <span className="spacer" />
        {durum?.bagli && (
          <>
            <button type="button" className="btn small secondary" onClick={yukle} disabled={yukleniyor}>
              <Icon name="refresh" size={14} /> Yenile
            </button>
            <button type="button" className="btn small secondary" onClick={() => { setAramaAcik(true); if (q) ara(q); }}>
              <Icon name="edit" size={14} /> Değiştir
            </button>
            <button type="button" className="btn small secondary" onClick={() => { if (confirm("Mikro cari bağlantısı kaldırılsın mı?")) esle(null); }} disabled={kaydediyor}>
              <Icon name="x" size={14} /> Kaldır
            </button>
          </>
        )}
      </div>

      {yukleniyor && !durum && <div className="skeleton" style={{ height: 70 }} />}

      {durum?.bagli && ozet && (
        <div className="cari-cards" style={{ marginTop: 4 }}>
          <div className={`cari-card ${ozet.bakiye > 0 ? "borc" : ""}`}>
            <span>Mikro Bakiye</span>
            <strong style={{ color: ozet.bakiye > 0 ? "var(--error)" : "var(--success)" }}>₺{fmt(ozet.bakiye)}</strong>
            <span className="muted" style={{ fontSize: 12 }}>{bakiyeAciklama(ozet.bakiye)}</span>
          </div>
          <div className={`cari-card ${ozet.vadesiGecen > 0 ? "borc" : ""}`}>
            <span>Vadesi Geçen (yaklaşık)</span>
            <strong style={{ color: ozet.vadesiGecen > 0 ? "var(--error)" : "inherit" }}>
              {ozet.vadeVar === false ? "—" : `₺${fmt(ozet.vadesiGecen)}`}
            </strong>
          </div>
          <div className="cari-card">
            <span>Toplam Borç</span>
            <strong>₺{fmt(ozet.borc)}</strong>
          </div>
          <div className="cari-card">
            <span>Toplam Alacak</span>
            <strong style={{ color: "var(--success)" }}>₺{fmt(ozet.alacak)}</strong>
          </div>
          <div className="cari-card">
            <span>Son Hareket</span>
            <strong>{ozet.sonHareket ? ozet.sonHareket.split("-").reverse().join(".") : "—"}</strong>
          </div>
        </div>
      )}

      {durum?.bagli && !ozet && !yukleniyor && (
        <div className="notice err">Mikro bakiyesi okunamadı: {durum.hata || "bilinmeyen hata"}</div>
      )}

      {durum && !durum.bagli && durum.ok && !aramaAcik && (
        <div className="notice info">
          Bu müşteri Mikro&apos;daki bir cari kartla henüz eşleştirilmedi.
          <button type="button" className="btn small" style={{ marginLeft: 10 }} onClick={() => { setAramaAcik(true); if (q) ara(q); }}>
            <Icon name="search" size={14} /> Mikro&apos;da bul
          </button>
        </div>
      )}

      {durum && !durum.ok && !durum.bagli && (
        <div className="notice err">{durum.hata || durum.error || "Mikro'ya ulaşılamadı."}</div>
      )}

      {aramaAcik && aramaKutusu}

      <p className="muted" style={{ fontSize: 12.5, marginTop: 10 }}>
        Bakiye Mikro&apos;dan canlı okunur (birkaç dakika önbellekte kalabilir); buradan Mikro&apos;ya kayıt yazılmaz.
        Pozitif bakiye müşterinin borcu demektir; tahsilatı buna göre isteyin.
      </p>
    </div>
  );
}
