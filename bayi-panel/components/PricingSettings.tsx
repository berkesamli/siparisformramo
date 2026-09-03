"use client";

// Bayi fiyat ayarları: çerçeve çarpanı, kur, paspartu/cam/baskı fiyatları,
// işçilik; firma bilgileri ve şifre değişikliği.

import { useEffect, useMemo, useState } from "react";
import type { DealerPricing } from "@/data/pricing";
import type { PublicDealer } from "@/lib/dealers";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// Örnek hesap: 50×70 cm eser, GC065 (1,50 USD/mt liste), 5 cm paspartu, düz cam
const SAMPLE = { usd: 1.5, code: "GC065", w: 500, h: 700, mat: 50 };

export default function PricingSettings() {
  const [pricing, setPricing] = useState<DealerPricing | null>(null);
  const [dealer, setDealer] = useState<PublicDealer | null>(null);
  const [autoRate, setAutoRate] = useState<number | null>(null);
  const [blob, setBlob] = useState(true);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);

  const [profile, setProfile] = useState({ name: "", contactName: "", phone: "", email: "", address: "", city: "", website: "" });
  const [pw, setPw] = useState({ current: "", next: "", again: "" });

  useEffect(() => {
    fetch("/api/ayarlar")
      .then((r) => r.json())
      .then((d) => {
        if (!d.ok) return;
        setPricing(d.pricing);
        setDealer(d.dealer);
        setAutoRate(d.autoRate);
        setBlob(d.blob !== false);
        setProfile({
          name: d.dealer.name || "",
          contactName: d.dealer.contactName || "",
          phone: d.dealer.phone || "",
          email: d.dealer.email || "",
          address: d.dealer.address || "",
          city: d.dealer.city || "",
          website: d.dealer.website || "",
        });
      })
      .catch(() => setMsg({ ok: false, text: "Ayarlar yüklenemedi." }));
  }, []);

  const effectiveRate = useMemo(() => {
    if (!pricing) return 0;
    if (pricing.usdRateMode === "manual" && pricing.usdRate > 0) return pricing.usdRate;
    return autoRate || pricing.usdRate || 0;
  }, [pricing, autoRate]);

  const sample = useMemo(() => {
    if (!pricing || !(effectiveRate > 0)) return null;
    const tlPerM = SAMPLE.usd * pricing.frameFactor * effectiveRate;
    const tw = SAMPLE.w + SAMPLE.mat * 2;
    const th = SAMPLE.h + SAMPLE.mat * 2;
    const perim = (2 * (tw + th)) / 1000 + 0.3;
    const area = (tw / 1000) * (th / 1000);
    const dk = pricing.mats.find((m) => m.code === "DK")?.price || 0;
    const cam = pricing.glasses.find((g) => g.name === "Düz Cam")?.price || 0;
    const frame = perim * tlPerM;
    const listCost = perim * SAMPLE.usd * effectiveRate;
    const mat = area * dk;
    const glass = area * cam;
    return { tlPerM, frame, listCost, mat, glass, total: frame + mat + glass + pricing.laborTL };
  }, [pricing, effectiveRate]);

  async function save(extra?: Record<string, unknown>) {
    if (!pricing) return;
    setSaving(true);
    setMsg(null);
    try {
      const res = await fetch("/api/ayarlar", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pricing, profile, ...extra }),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Kaydedilemedi");
      if (d.pricing) setPricing(d.pricing);
      if (d.dealer) setDealer(d.dealer);
      setMsg({ ok: true, text: "Ayarlar kaydedildi." });
      setPw({ current: "", next: "", again: "" });
    } catch (e: any) {
      setMsg({ ok: false, text: e.message || "Bir hata oluştu" });
    } finally {
      setSaving(false);
    }
  }

  function changePassword() {
    if (pw.next.length < 6) return setMsg({ ok: false, text: "Yeni şifre en az 6 karakter olmalı." });
    if (pw.next !== pw.again) return setMsg({ ok: false, text: "Yeni şifreler eşleşmiyor." });
    save({ currentPassword: pw.current, newPassword: pw.next });
  }

  if (!pricing || !dealer) return <p style={{ color: "var(--muted)" }}>Yükleniyor...</p>;

  const setNum = (v: string) => Math.max(0, parseFloat(String(v).replace(",", ".")) || 0);

  return (
    <div style={{ display: "grid", gap: 16 }}>
      {!blob && (
        <div className="notice err">Kalıcı depolama (Vercel Blob) yapılandırılmamış; ayarlar kaydedilemez.</div>
      )}
      {msg && <div className={`notice ${msg.ok ? "ok" : "err"}`}>{msg.text}</div>}

      {/* ---- Çerçeve & kur ---- */}
      <div className="card">
        <h2 style={{ marginTop: 0 }}>🖼️ Çerçeve Fiyatlandırması</h2>
        <p className="subtitle" style={{ marginTop: 0 }}>
          Çerçeve metre fiyatı = Olga toptan liste fiyatı (USD/mt) × <strong>çarpan</strong> × USD kuru.
          Çarpanı kendi kâr hedefinize göre belirleyin.
        </p>
        <div className="rw-grid2">
          <div>
            <label>Çerçeve Çarpanı</label>
            <input
              type="number" step="0.1" min="0.1"
              value={pricing.frameFactor}
              onChange={(e) => setPricing({ ...pricing, frameFactor: setNum(e.target.value) })}
            />
            <span style={{ fontSize: 12, color: "var(--muted)" }}>Olga perakende mağazası varsayılanı: 5</span>
          </div>
          <div>
            <label>USD Kuru</label>
            <div style={{ display: "flex", gap: 8 }}>
              <select
                style={{ width: 150 }}
                value={pricing.usdRateMode}
                onChange={(e) => setPricing({ ...pricing, usdRateMode: e.target.value as "auto" | "manual" })}
              >
                <option value="auto">Otomatik (TCMB)</option>
                <option value="manual">Sabit kur</option>
              </select>
              <input
                type="number" step="0.01" min="0"
                value={pricing.usdRate || ""}
                placeholder={autoRate ? `TCMB: ${autoRate}` : "örn. 47.50"}
                onChange={(e) => setPricing({ ...pricing, usdRate: setNum(e.target.value) })}
              />
            </div>
            <span style={{ fontSize: 12, color: "var(--muted)" }}>
              {autoRate ? `Bugünkü TCMB satış kuru: ${autoRate}` : "TCMB kuru alınamadı — sabit kur girin"}
              {pricing.usdRateMode === "auto" && " · Sabit kur, TCMB erişilemezse yedek olarak kullanılır."}
            </span>
          </div>
          <div>
            <label>İşçilik (kalem başına, ₺)</label>
            <input
              type="number" step="1" min="0"
              value={pricing.laborTL}
              onChange={(e) => setPricing({ ...pricing, laborTL: setNum(e.target.value) })}
            />
            <span style={{ fontSize: 12, color: "var(--muted)" }}>0 = işçilik ayrı yazılmaz (çarpanın içinde)</span>
          </div>
        </div>

        {sample && (
          <div className="notice info" style={{ marginTop: 14, fontSize: 13 }}>
            <strong>Örnek:</strong> 50×70 cm eser, {SAMPLE.code} profil ({SAMPLE.usd} $/mt), 5 cm düz karton paspartu, düz cam →
            çerçeve ₺{fmt(sample.frame)} (₺{fmt(sample.tlPerM)}/m; toptan liste maliyeti ≈ ₺{fmt(sample.listCost)}),
            paspartu ₺{fmt(sample.mat)}, cam ₺{fmt(sample.glass)}
            {pricing.laborTL > 0 && <>, işçilik ₺{fmt(pricing.laborTL)}</>} — <strong>toplam ₺{fmt(sample.total)}</strong>
          </div>
        )}
      </div>

      {/* ---- Paspartu / Cam / Baskı ---- */}
      <div className="rw-grid2" style={{ alignItems: "start" }}>
        <div className="card">
          <h2 style={{ marginTop: 0 }}>🎨 Paspartu (₺/m²)</h2>
          {pricing.mats.filter((m) => m.code !== "-").map((m, i) => (
            <div key={m.code} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <span style={{ flex: 1 }}>{m.icon} {m.name}</span>
              <input
                type="number" min="0" step="50" style={{ width: 130 }}
                value={m.price}
                onChange={(e) => {
                  const mats = pricing.mats.map((x) => (x.code === m.code ? { ...x, price: setNum(e.target.value) } : x));
                  setPricing({ ...pricing, mats });
                }}
              />
            </div>
          ))}
        </div>
        <div className="card">
          <h2 style={{ marginTop: 0 }}>🪟 Cam (₺/m²)</h2>
          {pricing.glasses.filter((g) => g.name !== "Cam Yok").map((g) => (
            <div key={g.name} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <span style={{ flex: 1 }}>{g.icon} {g.name}</span>
              <input
                type="number" min="0" step="50" style={{ width: 130 }}
                value={g.price}
                onChange={(e) => {
                  const glasses = pricing.glasses.map((x) => (x.name === g.name ? { ...x, price: setNum(e.target.value) } : x));
                  setPricing({ ...pricing, glasses });
                }}
              />
            </div>
          ))}
          <h2 style={{ marginTop: 18 }}>🖨️ Baskı ($/m²)</h2>
          {pricing.prints.filter((p) => p.name !== "Baskı Yok").map((p) => (
            <div key={p.name} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <span style={{ flex: 1 }}>{p.icon} {p.name}</span>
              <input
                type="number" min="0" step="0.1" style={{ width: 130 }}
                value={p.usdPerM2}
                onChange={(e) => {
                  const prints = pricing.prints.map((x) => (x.name === p.name ? { ...x, usdPerM2: setNum(e.target.value) } : x));
                  setPricing({ ...pricing, prints });
                }}
              />
            </div>
          ))}
        </div>
      </div>

      {/* ---- Firma bilgileri ---- */}
      <div className="card">
        <h2 style={{ marginTop: 0 }}>🏪 Firma Bilgileri</h2>
        <p className="subtitle" style={{ marginTop: 0 }}>PDF, fiş ve WhatsApp mesajlarında müşteriniz bu bilgileri görür.</p>
        <div className="rw-grid2">
          <div><label>Firma Adı *</label><input value={profile.name} onChange={(e) => setProfile({ ...profile, name: e.target.value })} /></div>
          <div><label>Yetkili</label><input value={profile.contactName} onChange={(e) => setProfile({ ...profile, contactName: e.target.value })} /></div>
          <div><label>Telefon *</label><input value={profile.phone} onChange={(e) => setProfile({ ...profile, phone: e.target.value })} /></div>
          <div><label>E-posta</label><input type="email" value={profile.email} onChange={(e) => setProfile({ ...profile, email: e.target.value })} /></div>
          <div><label>Şehir</label><input value={profile.city} onChange={(e) => setProfile({ ...profile, city: e.target.value })} /></div>
          <div><label>Web sitesi / Instagram</label><input value={profile.website} onChange={(e) => setProfile({ ...profile, website: e.target.value })} placeholder="www.firmam.com" /></div>
          <div style={{ gridColumn: "1 / -1" }}><label>Adres</label><input value={profile.address} onChange={(e) => setProfile({ ...profile, address: e.target.value })} /></div>
        </div>
        <div style={{ fontSize: 12.5, color: "var(--muted)", marginTop: 10 }}>
          Bayi kodu: <strong>{dealer.slug}</strong> · Kullanıcı adı: <strong>{dealer.username}</strong>
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <button className="btn" disabled={saving || !blob} onClick={() => save()}>
          {saving ? "Kaydediliyor..." : "💾 Ayarları Kaydet"}
        </button>
      </div>

      {/* ---- Şifre ---- */}
      <div className="card">
        <h2 style={{ marginTop: 0 }}>🔑 Şifre Değiştir</h2>
        <div className="rw-grid2">
          <div><label>Mevcut Şifre</label><input type="password" value={pw.current} onChange={(e) => setPw({ ...pw, current: e.target.value })} /></div>
          <div />
          <div><label>Yeni Şifre</label><input type="password" value={pw.next} onChange={(e) => setPw({ ...pw, next: e.target.value })} /></div>
          <div><label>Yeni Şifre (tekrar)</label><input type="password" value={pw.again} onChange={(e) => setPw({ ...pw, again: e.target.value })} /></div>
        </div>
        <button className="btn secondary" style={{ marginTop: 12 }} disabled={saving || !pw.current || !pw.next} onClick={changePassword}>
          Şifreyi Güncelle
        </button>
      </div>
    </div>
  );
}
