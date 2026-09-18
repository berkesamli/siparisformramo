"use client";

// Sahiplere özel: Mikro Jump API bağlantı durumu, bağlantı denemesi ve cari bakiye sorgusu.

import { useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Durum {
  kurulu: boolean;
  url: string | null;
  firma: string | null;
  kullanici: string | null;
  yil: string;
  apiKey: boolean;
  sifre: boolean;
}
interface CariSatir { cari_kod: string; unvan: string }
interface CariOzet {
  cariKod: string; unvan: string; borc: number; alacak: number; bakiye: number; vadesiGecen: number; sonHareket: string | null;
}

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export default function MikroAyarlari() {
  const [durum, setDurum] = useState<Durum | null>(null);
  const [hata, setHata] = useState("");
  const [deniyor, setDeniyor] = useState(false);
  const [cariler, setCariler] = useState<CariSatir[] | null>(null);
  const [testHata, setTestHata] = useState("");
  const [testRaw, setTestRaw] = useState("");
  const [cariKod, setCariKod] = useState("");
  const [sorguluyor, setSorguluyor] = useState(false);
  const [ozet, setOzet] = useState<CariOzet | null>(null);
  const [ozetHata, setOzetHata] = useState("");

  useEffect(() => {
    fetch("/api/mikro/test")
      .then((r) => r.json())
      .then((d) => (d.ok ? setDurum(d) : setHata(d.error || "Durum alınamadı.")))
      .catch(() => setHata("Sunucuya ulaşılamadı."));
  }, []);

  async function baglantiDene() {
    setDeniyor(true); setCariler(null); setTestHata(""); setTestRaw("");
    try {
      const r = await fetch("/api/mikro/test", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ islem: "baglanti" }) });
      const d = await r.json();
      if (d.ok) setCariler(d.cariler || []);
      else { setTestHata(d.hata || d.error || "Bağlantı başarısız."); if (d.raw) setTestRaw(JSON.stringify(d.raw).slice(0, 600)); }
    } catch { setTestHata("Sunucuya ulaşılamadı."); }
    finally { setDeniyor(false); }
  }

  async function cariSorgula() {
    if (!cariKod.trim()) return;
    setSorguluyor(true); setOzet(null); setOzetHata("");
    try {
      const r = await fetch("/api/mikro/test", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ islem: "cari", cariKod: cariKod.trim() }) });
      const d = await r.json();
      if (d.ok && d.ozet) setOzet(d.ozet);
      else setOzetHata(d.hata || d.error || "Sorgu başarısız.");
    } catch { setOzetHata("Sunucuya ulaşılamadı."); }
    finally { setSorguluyor(false); }
  }

  const Satir = ({ ok, baslik, aciklama }: { ok: boolean; baslik: string; aciklama: string }) => (
    <li className="row" style={{ padding: "10px 0", borderBottom: "1px solid var(--border)", alignItems: "flex-start", flexWrap: "nowrap" }}>
      <span className={`badge ${ok ? "ok" : "warn"}`} style={{ marginTop: 2 }}>
        <Icon name={ok ? "check-circle" : "alert"} size={13} /> {ok ? "Hazır" : "Eksik"}
      </span>
      <span style={{ minWidth: 0 }}>
        <strong style={{ display: "block", fontSize: 14 }}>{baslik}</strong>
        <span className="muted" style={{ fontSize: 12.5 }}>{aciklama}</span>
      </span>
    </li>
  );

  return (
    <div className="card">
      <div className="card-head">
        <span className="card-head-icon"><Icon name="dollar" size={18} /></span>
        <div>
          <h2>Mikro Bağlantısı</h2>
          <span className="card-head-sub">Ankara sunucusundaki Mikro Jump 17 API&apos;si · yalnızca okuma (cari bakiye, vade)</span>
        </div>
        <span className="spacer" />
        <button type="button" className="btn" onClick={baglantiDene} disabled={!durum?.kurulu || deniyor}>
          <Icon name="refresh" size={16} /> {deniyor ? "Deneniyor…" : "Bağlantıyı dene"}
        </button>
      </div>
      {hata && <div className="notice err">{hata}</div>}
      {!durum && !hata && <div className="skeleton" style={{ height: 120 }} />}
      {durum && (
        <ul style={{ listStyle: "none" }}>
          <Satir ok={!!durum.url} baslik="Sunucu adresi (MIKRO_API_URL)" aciklama={durum.url || "Tanımlı değil. Örn. https://olgaserver.tailbbb7b8.ts.net"} />
          <Satir ok={durum.apiKey} baslik="API anahtarı (MIKRO_API_KEY)" aciklama={durum.apiKey ? "Tanımlı." : "Mikro'nun verdiği anahtar."} />
          <Satir ok={!!durum.firma && !!durum.kullanici && durum.sifre} baslik="Firma / kullanıcı / şifre" aciklama={`Veri tabanı: ${durum.firma || "—"} · Kullanıcı: ${durum.kullanici || "—"} · Şifre: ${durum.sifre ? "tanımlı" : "eksik"} · Çalışma yılı: ${durum.yil}`} />
        </ul>
      )}
      {testHata && (
        <div className="notice err" style={{ marginTop: 14 }}>
          <strong>Bağlantı başarısız</strong> · {testHata}
          {testRaw && <pre style={{ whiteSpace: "pre-wrap", fontSize: 12, marginTop: 6 }}>{testRaw}</pre>}
        </div>
      )}
      {cariler && (
        <div className="notice ok" style={{ marginTop: 14 }}>
          <strong>Bağlantı çalışıyor.</strong> Mikro&apos;dan gelen ilk cari kartlar:
          <ul style={{ margin: "6px 0 0 18px" }}>
            {cariler.map((c) => <li key={c.cari_kod}><code>{c.cari_kod}</code> · {c.unvan}</li>)}
            {cariler.length === 0 && <li>Liste boş döndü.</li>}
          </ul>
        </div>
      )}

      <div className="row" style={{ marginTop: 16, gap: 8 }}>
        <input
          style={{ maxWidth: 220 }}
          placeholder="Cari kodu, örn. 120.01.001"
          value={cariKod}
          onChange={(e) => setCariKod(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") cariSorgula(); }}
        />
        <button type="button" className="btn secondary" onClick={cariSorgula} disabled={!durum?.kurulu || sorguluyor || !cariKod.trim()}>
          <Icon name="search" size={16} /> {sorguluyor ? "Sorgulanıyor…" : "Cari bakiye sorgula"}
        </button>
      </div>
      {ozetHata && <div className="notice err" style={{ marginTop: 12 }}>{ozetHata}</div>}
      {ozet && (
        <div className="notice info" style={{ marginTop: 12 }}>
          <strong>{ozet.unvan || ozet.cariKod}</strong> <span className="muted">({ozet.cariKod})</span>
          <div className="grid cols-2" style={{ marginTop: 8, gap: 6 }}>
            <div>Bakiye: <strong>₺ {fmt(ozet.bakiye)}</strong> {ozet.bakiye > 0 ? "(müşteri borçlu)" : ozet.bakiye < 0 ? "(müşteri alacaklı)" : ""}</div>
            <div>Vadesi geçen (yaklaşık): <strong style={{ color: ozet.vadesiGecen > 0 ? "var(--error)" : "inherit" }}>₺ {fmt(ozet.vadesiGecen)}</strong></div>
            <div>Toplam borç: ₺ {fmt(ozet.borc)}</div>
            <div>Toplam alacak: ₺ {fmt(ozet.alacak)}</div>
            <div>Son hareket: {ozet.sonHareket || "—"}</div>
          </div>
        </div>
      )}
      <p className="muted" style={{ fontSize: 12.5, marginTop: 12 }}>
        Bu kart yalnızca okuma yapar; Mikro&apos;ya hiçbir kayıt yazılmaz. Bağlantı doğrulanınca müşteri kartı ve sipariş formuna bakiye uyarısı eklenecek.
      </p>
    </div>
  );
}
