"use client";

// Günlük kur belirleme — yalnızca firma sahipleri (Berke, Özgür, Gültekin).
// Sabah kur girilir, gün boyu bütün sipariş formlarına otomatik gelir ve
// diğer çalışanlar değiştiremez; herkes aynı kurdan sipariş girer.

import { useCallback, useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";
import { sayi } from "@/lib/num";

interface Rates {
  rate: number;
  euroRate: number;
  updatedAt: string;
  by: string;
  sabit?: boolean;
}

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 4,
  });

// Kur değeri satır sonunda bölünmesin ("Dolar: ₺ 41,85" tek parça kalır).
const NOWRAP = { whiteSpace: "nowrap" } as const;

export default function GunlukKur() {
  const [mevcut, setMevcut] = useState<Rates | null>(null);
  const [yuklendi, setYuklendi] = useState(false);
  const [usd, setUsd] = useState("");
  const [eur, setEur] = useState("");
  const [kaydediliyor, setKaydediliyor] = useState(false);
  const [msg, setMsg] = useState("");
  const [err, setErr] = useState("");

  const yukle = useCallback(async () => {
    try {
      const res = await fetch("/api/rates");
      const d = await res.json();
      if (d.ok) {
        setMevcut(d.rates || null);
        if (d.rates?.rate > 0) setUsd(String(d.rates.rate));
        if (d.rates?.euroRate > 0) setEur(String(d.rates.euroRate));
      }
    } catch {
      setErr("Kur bilgisi alınamadı.");
    } finally {
      setYuklendi(true);
    }
  }, []);

  useEffect(() => {
    yukle();
  }, [yukle]);

  async function kaydet() {
    const rate = sayi(usd);
    const euroRate = sayi(eur);
    if (rate <= 0 && euroRate <= 0) {
      setErr("En az bir kur değeri girin.");
      return;
    }
    setKaydediliyor(true);
    setErr("");
    setMsg("");
    try {
      const res = await fetch("/api/rates", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rate, euroRate }),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Kaydedilemedi");
      setMevcut(d.rates);
      setMsg("Günün kuru kaydedildi — tüm sipariş formlarına bu kur gelecek.");
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Bir hata oluştu.");
    } finally {
      setKaydediliyor(false);
    }
  }

  const bugun = new Date().toLocaleDateString("tr-TR", {
    day: "numeric",
    month: "long",
    year: "numeric",
    timeZone: "Europe/Istanbul",
  });

  return (
    <div className="card">
      <div className="card-head">
        <span className="card-head-icon"><Icon name="dollar" size={18} /></span>
        <div><h2>Bugünün Kuru</h2></div>
        <span className="spacer" />
        <div className="card-head-actions">
          <span className="badge">{bugun}</span>
        </div>
      </div>

      {yuklendi && mevcut && (mevcut.rate > 0 || mevcut.euroRate > 0) ? (
        <div className="notice ok">
          <div className="row" style={{ gap: "4px 20px", fontSize: 15 }}>
            {mevcut.rate > 0 && <b style={NOWRAP}>Dolar: ₺ {fmt(mevcut.rate)}</b>}
            {mevcut.euroRate > 0 && <b style={NOWRAP}>Euro: ₺ {fmt(mevcut.euroRate)}</b>}
          </div>
          <div className="small" style={{ marginTop: 4 }}>
            {mevcut.sabit ? "Yetkili tarafından belirlendi" : "İlk siparişten alındı"} —{" "}
            {mevcut.by} ·{" "}
            {new Date(mevcut.updatedAt).toLocaleTimeString("tr-TR", {
              hour: "2-digit",
              minute: "2-digit",
              timeZone: "Europe/Istanbul",
            })}
          </div>
        </div>
      ) : yuklendi ? (
        <div className="notice warn">
          ⚠️ Bugün için kur henüz girilmedi. Kur girilene kadar çalışanlar
          sipariş formunda kuru kendileri yazmak zorunda kalır.
        </div>
      ) : (
        <p className="text-2" style={{ marginTop: 14 }}>Yükleniyor…</p>
      )}

      <div className="field-row" style={{ marginTop: 16 }}>
        <div className="field">
          <label>Dolar Kuru (TL/USD)</label>
          <input
            type="text"
            inputMode="decimal"
            value={usd}
            onChange={(e) => setUsd(e.target.value)}
            placeholder="örn. 41,85"
          />
        </div>
        <div className="field">
          <label>Euro Kuru (TL/EUR)</label>
          <input
            type="text"
            inputMode="decimal"
            value={eur}
            onChange={(e) => setEur(e.target.value)}
            placeholder="örn. 48,60"
          />
        </div>
      </div>

      <div className="row">
        <button className="btn" disabled={kaydediliyor} onClick={kaydet}>
          {kaydediliyor ? "Kaydediliyor…" : mevcut?.sabit ? "Kuru Güncelle" : "Günün Kurunu Belirle"}
        </button>
        <button className="btn secondary" onClick={yukle} disabled={kaydediliyor}>
          <Icon name="refresh" size={14} /> Yenile
        </button>
      </div>

      {msg && <div className="notice ok">{msg}</div>}
      {err && <div className="notice err">{err}</div>}

      <p className="muted small" style={{ marginTop: 16 }}>
        💡 Kuru siz belirledikten sonra sipariş formundaki kur alanı diğer
        çalışanlarda kilitlenir — herkes bu kurdan sipariş girer. Gün içinde
        değiştirirseniz yeni siparişler yeni kuru kullanır; daha önce alınan
        siparişler kendi kurunu korur.
      </p>
    </div>
  );
}
