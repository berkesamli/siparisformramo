"use client";

// Yeni sipariş bildirimi — üst çubuktaki zil. Panel açıkken sipariş sayaçlarını
// izler; sayı artınca zil sesi çalar, tarayıcı bildirimi ve ekranda uyarı gösterir.
// Maliyet notu: her yoklama Blob'dan yalnızca iki küçük sayaç dosyası okur ve
// sekme görünür değilken hiç yoklama yapılmaz. (Eski NewOrderAlert'in devamı.)

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import Icon from "./Icon";
import { DUYURU_EVENT, isSeen, markSeen } from "./AnnouncementModal";
import type { Duyuru } from "./types";

const ARALIK_MS = 90_000;
const LS_ACIK = "orderAlertOn";
const LS_SON = "orderAlertSeen"; // "toptan:perakende"

function zilCal() {
  try {
    const Ctx = window.AudioContext || (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
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
    setTimeout(() => ctx.close().catch(() => {}), 1500);
  } catch { /* ses çalınamazsa görsel uyarı yeterli */ }
}

interface Uyari { id: number; metin: string; href: string; zaman: string; goruldu: boolean; }

/** Uzun duyuru metnini kısaltır (tam metin title özniteliğinde kalır). */
function kisalt(metin: string, uzunluk = 90): string {
  if (metin.length <= uzunluk) return metin;
  return metin.slice(0, uzunluk).trimEnd() + "…";
}

const BOS_DUYURU: Duyuru[] = [];

export default function NotificationBell({
  onStatsDirty,
  duyurular = BOS_DUYURU,
  username = "",
}: {
  onStatsDirty?: () => void;
  duyurular?: Duyuru[];
  username?: string;
}) {
  const [acik, setAcik] = useState(false);
  const [uyarilar, setUyarilar] = useState<Uyari[]>([]);
  const [toast, setToast] = useState<Uyari | null>(null);
  const [menu, setMenu] = useState(false);
  // Duyurular: görülmemiş id'ler (rozet + kırmızı nokta). Menü açılırken görülmüş
  // sayılır ama "Yeni" rozeti menü kapanana dek kalsın diye anlık görüntü tutulur.
  const [duyuruYeni, setDuyuruYeni] = useState<Set<string>>(new Set());
  const [menuYeni, setMenuYeni] = useState<Set<string>>(new Set());
  const sonRef = useRef<{ t: number; p: number } | null>(null);
  const acikRef = useRef(false);
  const ref = useRef<HTMLDivElement>(null);
  acikRef.current = acik;

  useEffect(() => {
    try {
      setAcik(localStorage.getItem(LS_ACIK) === "1");
      const raw = localStorage.getItem(LS_SON);
      if (raw) {
        const [t, p] = raw.split(":").map((x) => Number(x) || 0);
        sonRef.current = { t, p };
      }
    } catch { /* depolama yok */ }
  }, []);

  // Görülmemiş duyurular — depodan okunur; modal veya başka bir yerde işaretlenince yenilenir
  useEffect(() => {
    const hesapla = () => setDuyuruYeni(new Set(duyurular.filter((d) => !isSeen(d.id, username)).map((d) => d.id)));
    hesapla();
    window.addEventListener(DUYURU_EVENT, hesapla);
    window.addEventListener("storage", hesapla);
    return () => {
      window.removeEventListener(DUYURU_EVENT, hesapla);
      window.removeEventListener("storage", hesapla);
    };
  }, [duyurular, username]);

  const yokla = useCallback(async () => {
    if (document.visibilityState !== "visible") return;
    try {
      const r = await fetch("/api/orders/latest");
      const d = await r.json();
      if (!d?.ok) return;
      const simdi = { t: Number(d.toptan) || 0, p: Number(d.perakende) || 0 };
      const son = sonRef.current;
      sonRef.current = simdi;
      try { localStorage.setItem(LS_SON, `${simdi.t}:${simdi.p}`); } catch { /* yok */ }
      if (!son) return;
      const yeniToptan = simdi.t - son.t;
      const yeniPerakende = simdi.p - son.p;
      if (yeniToptan <= 0 && yeniPerakende <= 0) return;
      onStatsDirty?.();
      if (!acikRef.current) return;
      const parca: string[] = [];
      if (yeniToptan > 0) parca.push(`${yeniToptan} yeni toptan sipariş`);
      if (yeniPerakende > 0) parca.push(`${yeniPerakende} yeni perakende sipariş`);
      const metin = parca.join(", ") + "!";
      const href = yeniToptan > 0 ? "/panel/siparisler" : "/panel/perakende/siparisler";
      const u: Uyari = {
        id: Date.now(),
        metin,
        href,
        zaman: new Date().toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit" }),
        goruldu: false,
      };
      setUyarilar((l) => [u, ...l].slice(0, 12));
      setToast(u);
      zilCal();
      if ("Notification" in window && Notification.permission === "granted") {
        try { new Notification("Olga Çerçeve — Yeni Sipariş", { body: metin }); } catch { /* engelli */ }
      }
    } catch { /* ağ hatasında sessiz kal */ }
  }, [onStatsDirty]);

  useEffect(() => {
    yokla();
    const id = setInterval(yokla, ARALIK_MS);
    const gorunurluk = () => { if (document.visibilityState === "visible") yokla(); };
    document.addEventListener("visibilitychange", gorunurluk);
    return () => {
      clearInterval(id);
      document.removeEventListener("visibilitychange", gorunurluk);
    };
  }, [yokla]);

  useEffect(() => {
    if (!menu) return;
    const onDoc = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setMenu(false); };
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") setMenu(false); };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => { document.removeEventListener("mousedown", onDoc); document.removeEventListener("keydown", onKey); };
  }, [menu]);

  // Toast 12 sn sonra kendiliğinden kapanır
  useEffect(() => {
    if (!toast) return;
    const id = setTimeout(() => setToast(null), 12_000);
    return () => clearTimeout(id);
  }, [toast]);

  function toggle() {
    const yeni = !acik;
    setAcik(yeni);
    try { localStorage.setItem(LS_ACIK, yeni ? "1" : "0"); } catch { /* yok */ }
    if (yeni) {
      zilCal();
      if ("Notification" in window && Notification.permission === "default") {
        Notification.requestPermission().catch(() => {});
      }
    }
  }

  function openMenu() {
    const acilacak = !menu;
    setMenu(acilacak);
    setUyarilar((l) => l.map((u) => ({ ...u, goruldu: true })));
    if (acilacak) {
      // Menü açılınca duyurular görülmüş sayılır; rozet bu açılış boyunca görünür kalır
      setMenuYeni(new Set(duyuruYeni));
      duyuruYeni.forEach((id) => markSeen(id, username));
    } else {
      setMenuYeni(new Set());
    }
  }

  const okunmamis = uyarilar.some((u) => !u.goruldu) || duyuruYeni.size > 0;

  return (
    <>
      <button
        type="button"
        className={`tb-live ${acik ? "" : "off"}`}
        onClick={toggle}
        title={acik ? "Yeni sipariş bildirimi açık — kapatmak için tıklayın" : "Yeni sipariş bildirimi kapalı — açmak için tıklayın"}
      >
        <span className="dot" /> {acik ? "Canlı" : "Sessiz"}
      </button>
      <div className="menu-wrap" ref={ref}>
        <button type="button" className="btn icon ghost nb-btn" onClick={openMenu} aria-label="Bildirimler" aria-expanded={menu} title="Bildirimler">
          <Icon name={acik ? "bell" : "bell-off"} size={19} />
          {okunmamis && <span className="nb-dot" />}
        </button>
        {menu && (
          <div className="menu" role="menu" style={{ minWidth: 300 }}>
            <div className="menu-head">
              <strong>Bildirimler</strong>
              <span>Yeni sipariş düştüğünde ses + ekran uyarısı</span>
            </div>
            {duyurular.length > 0 && (
              <>
                <div className="nb-sec">Duyurular</div>
                <div className="nb-list">
                  {duyurular.map((d) => {
                    const yeni = menuYeni.has(d.id);
                    const icerik = (
                      <>
                        <span className="nb-item-icon"><Icon name={d.ikon || "sparkles"} size={16} /></span>
                        <span className="nb-item-main">
                          <strong>{d.baslik}{yeni && <span className="badge brand">Yeni</span>}</strong>
                          <span className="nb-ann-text">{kisalt(d.metin)}</span>
                        </span>
                      </>
                    );
                    const tikla = () => { markSeen(d.id, username); setMenu(false); };
                    return d.href ? (
                      <Link key={d.id} href={d.href} className="nb-item nb-ann" title={d.metin} onClick={tikla}>{icerik}</Link>
                    ) : (
                      <div key={d.id} className="nb-item nb-ann" title={d.metin}>{icerik}</div>
                    );
                  })}
                </div>
                <div className="menu-sep" />
              </>
            )}
            <div className="nb-toggle">
              <Icon name={acik ? "bell" : "bell-off"} size={16} />
              <span>Sesli / tarayıcı bildirimi</span>
              <span className="spacer" />
              <button type="button" className={`switch ${acik ? "on" : ""}`} onClick={toggle} aria-pressed={acik} aria-label="Bildirimi aç/kapat" />
            </div>
            <div className="menu-sep" />
            <div className="nb-list">
              {uyarilar.length === 0 ? (
                <div className="menu-note">Bu oturumda yeni sipariş uyarısı yok. Sayaçlar 90 saniyede bir kontrol edilir.</div>
              ) : (
                uyarilar.map((u) => (
                  <Link key={u.id} href={u.href} className="nb-item" onClick={() => setMenu(false)}>
                    <span className="nb-item-icon"><Icon name="inbox" size={16} /></span>
                    <span>
                      <strong>{u.metin}</strong>
                      <span>{u.zaman} · Siparişlere git</span>
                    </span>
                  </Link>
                ))
              )}
            </div>
          </div>
        )}
      </div>
      {toast && (
        <div className="toast-wrap" role="status" aria-live="polite">
          <div className="toast">
            <span className="toast-icon"><Icon name="inbox" size={18} /></span>
            <div className="toast-body">
              <strong>🛎 {toast.metin}</strong>
              <div className="toast-actions">
                <Link className="btn small" href={toast.href} onClick={() => setToast(null)}>Siparişlere Git</Link>
                <button type="button" className="btn small secondary" onClick={() => setToast(null)}>Kapat</button>
              </div>
            </div>
            <button type="button" className="btn icon ghost small" onClick={() => setToast(null)} aria-label="Kapat"><Icon name="x" size={16} /></button>
          </div>
        </div>
      )}
    </>
  );
}
