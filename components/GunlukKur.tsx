"use client";

// Günlük kur belirleme — yalnızca firma sahipleri (Berke, Özgür, Gültekin).
// Sabah kur girilir, gün boyu bütün sipariş formlarına otomatik gelir ve
// diğer çalışanlar değiştiremez; herkes aynı kurdan sipariş girer.

import { useCallback, useEffect, useState } from "react";
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
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <h2 style={{ margin: 0, fontSize: 17 }}>Bugünün Kuru</h2>
        <span className="badge">{bugun}</span>
      </div>

      {yuklendi && mevcut && (mevcut.rate > 0 || mevcut.euroRate > 0) ? (
        <div className="notice ok" style={{ marginTop: 14 }}>
          <div style={{ fontSize: 15 }}>
            {mevcut.rate > 0 && (
              <>
                <b>Dolar: ₺ {fmt(mevcut.rate)}</b>
                {mevcut.euroRate > 0 ? "  ·  " : ""}
              </>
            )}
            {mevcut.euroRate > 0 && <b>Euro: ₺ {fmt(mevcut.euroRate)}</b>}
          </div>
          <div style={{ fontSize: 12.5, marginTop: 4 }}>
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
        <div className="notice warn" style={{ marginTop: 14 }}>
          ⚠️ Bugün için kur henüz girilmedi. Kur girilene kadar çalışanlar
          sipariş formunda kuru kendileri yazmak zorunda kalır.
        </div>
      ) : (
        <p style={{ color: "var(--text-2)", marginTop: 14 }}>Yükleniyor…</p>
      )}

      <div
        className="grid"
        style={{ gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", marginTop: 16 }}
      >
        <div>
          <label>Dolar Kuru (TL/USD)</label>
          <input
            type="text"
            
            inputMode="decimal"
            value={usd}
            onChange={(e) => setUsd(e.target.value)}
            placeholder="örn. 41,85"
          />
        </div>
        <div>
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

      <div style={{ display: "flex", gap: 10, marginTop: 16, flexWrap: "wrap" }}>
        <button className="btn" disabled={kaydediliyor} onClick={kaydet}>
          {kaydediliyor ? "Kaydediliyor…" : mevcut?.sabit ? "Kuru Güncelle" : "Günün Kurunu Belirle"}
        </button>
        <button className="btn secondary" onClick={yukle} disabled={kaydediliyor}>
          ↻ Yenile
        </button>
      </div>

      {msg && <div className="notice ok" style={{ marginTop: 12 }}>{msg}</div>}
      {err && <div className="notice err" style={{ marginTop: 12 }}>{err}</div>}

      <p style={{ color: "var(--muted)", fontSize: 12.5, marginTop: 16 }}>
        💡 Kuru siz belirledikten sonra sipariş formundaki kur alanı diğer
        çalışanlarda kilitlenir — herkes bu kurdan sipariş girer. Gün içinde
        değiştirirseniz yeni siparişler yeni kuru kullanır; daha önce alınan
        siparişler kendi kurunu korur.
      </p>
    </div>
  );
}
