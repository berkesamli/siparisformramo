"use client";

// Yeni sipariş uyarısı — panel açıkken sipariş sayaçlarını izler; sayı artınca
// zil sesi çalar, tarayıcı bildirimi ve ekranda uyarı gösterir. Amaç: personelin
// siparişler ekranını sürekli elle kontrol etmek zorunda kalmaması.
//
// Maliyet notu: her yoklama Blob'dan yalnızca iki küçük sayaç dosyası okur ve
// sekme görünür değilken hiç yoklama yapılmaz.

import { useCallback, useEffect, useRef, useState } from "react";

const ARALIK_MS = 90_000; // yoklama aralığı
const LS_ACIK = "orderAlertOn";
const LS_SON = "orderAlertSeen"; // "toptan:perakende"

// Zil sesi — dosya gerektirmesin diye WebAudio ile üretilir (iki kısa ding).
function zilCal() {
  try {
    const Ctx =
      window.AudioContext ||
      (window as unknown as { webkitAudioContext: typeof AudioContext })
        .webkitAudioContext;
    const ctx = new Ctx();
    const ding = (t: number, hz: number) => {
      const o = ctx.createOscillator();
      const g = ctx.createGain();
      o.type = "sine";
      o.frequency.value = hz;
      g.gain.setValueAtTime(0.0001, t);
      g.gain.exponentialRampToValueAtTime(0.35, t + 0.02);
      g.gain.exponentialRampToValueAtTime(0.0001, t + 0.7);
      o.connect(g).connect(ctx.destination);
      o.start(t);
      o.stop(t + 0.75);
    };
    ding(ctx.currentTime, 880);
    ding(ctx.currentTime + 0.25, 1175);
    // Bağlamı açık bırakmayalım — bazı tarayıcılar sekme başına sınır koyar.
    setTimeout(() => ctx.close().catch(() => {}), 1500);
  } catch {
    /* ses çalınamazsa görsel uyarı yeterli */
  }
}

interface Uyari {
  metin: string;
  href: string;
}

export default function NewOrderAlert() {
  const [acik, setAcik] = useState(false);
  const [uyari, setUyari] = useState<Uyari | null>(null);
  const sonRef = useRef<{ t: number; p: number } | null>(null);
  const acikRef = useRef(false);
  acikRef.current = acik;

  useEffect(() => {
    setAcik(localStorage.getItem(LS_ACIK) === "1");
    const raw = localStorage.getItem(LS_SON);
    if (raw) {
      const [t, p] = raw.split(":").map((x) => Number(x) || 0);
      sonRef.current = { t, p };
    }
  }, []);

  const yokla = useCallback(async () => {
    if (document.visibilityState !== "visible") return;
    try {
      const r = await fetch("/api/orders/latest");
      const d = await r.json();
      if (!d?.ok) return;
      const simdi = { t: Number(d.toptan) || 0, p: Number(d.perakende) || 0 };
      const son = sonRef.current;
      sonRef.current = simdi;
      localStorage.setItem(LS_SON, `${simdi.t}:${simdi.p}`);
      if (!son) return; // ilk yoklama — kıyas noktası yok

      const yeniToptan = simdi.t - son.t;
      const yeniPerakende = simdi.p - son.p;
      if (yeniToptan <= 0 && yeniPerakende <= 0) return;
      if (!acikRef.current) return;

      const parca: string[] = [];
      if (yeniToptan > 0) parca.push(`${yeniToptan} yeni toptan sipariş`);
      if (yeniPerakende > 0) parca.push(`${yeniPerakende} yeni perakende sipariş`);
      const metin = parca.join(", ") + "!";
      const href =
        yeniToptan > 0 ? "/panel/siparisler" : "/panel/perakende/siparisler";

      setUyari({ metin, href });
      zilCal();
      if ("Notification" in window && Notification.permission === "granted") {
        try {
          new Notification("Olga Çerçeve — Yeni Sipariş", { body: metin });
        } catch {
          /* bildirimi engelleyen tarayıcıda ses + ekran uyarısı yeterli */
        }
      }
    } catch {
      /* ağ hatasında sessiz kal; sonraki yoklamada tekrar denenir */
    }
  }, []);

  useEffect(() => {
    yokla();
    const id = setInterval(yokla, ARALIK_MS);
    const gorunurluk = () => {
      if (document.visibilityState === "visible") yokla();
    };
    document.addEventListener("visibilitychange", gorunurluk);
    return () => {
      clearInterval(id);
      document.removeEventListener("visibilitychange", gorunurluk);
    };
  }, [yokla]);

  function toggle() {
    const yeni = !acik;
    setAcik(yeni);
    localStorage.setItem(LS_ACIK, yeni ? "1" : "0");
    if (yeni) {
      // Buton tıklaması bir kullanıcı hareketi olduğu için tarayıcı sese ve
      // bildirim iznine burada izin verir — sonrası için kapıyı açıyoruz.
      zilCal();
      if ("Notification" in window && Notification.permission === "default") {
        Notification.requestPermission().catch(() => {});
      }
    }
  }

  return (
    <>
      <button
        onClick={toggle}
        title={
          acik
            ? "Yeni sipariş bildirimi açık — kapatmak için tıklayın"
            : "Yeni sipariş bildirimi kapalı — açmak için tıklayın"
        }
        aria-label="Yeni sipariş bildirimi"
        style={{
          position: "fixed",
          bottom: 18,
          left: 18,
          zIndex: 60,
          width: 44,
          height: 44,
          borderRadius: "50%",
          border: "1px solid var(--border, #444)",
          background: acik ? "var(--brand, #7c5cff)" : "var(--card, #1c1c22)",
          color: "#fff",
          fontSize: 20,
          cursor: "pointer",
          opacity: acik ? 1 : 0.65,
        }}
      >
        {acik ? "🔔" : "🔕"}
      </button>

      {uyari && (
        <div
          style={{
            position: "fixed",
            bottom: 74,
            left: 18,
            zIndex: 60,
            maxWidth: 300,
            background: "var(--card, #1c1c22)",
            border: "1px solid var(--brand, #7c5cff)",
            borderRadius: 10,
            padding: "12px 14px",
            boxShadow: "0 6px 24px rgba(0,0,0,.35)",
          }}
        >
          <strong>🛎 {uyari.metin}</strong>
          <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
            <a className="btn small" href={uyari.href}>
              Siparişlere Git
            </a>
            <button className="btn small secondary" onClick={() => setUyari(null)}>
              Kapat
            </button>
          </div>
        </div>
      )}
    </>
  );
}
