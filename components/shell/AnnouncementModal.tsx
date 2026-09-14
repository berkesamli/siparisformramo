"use client";

// Duyuru modali — girişte, henüz görülmemiş ilk (en yeni) duyuruyu bir kez gösterir.
// Durum tarayıcıda kullanıcı adına göre saklanır (olga-duyuru:<id>:<kullanıcı>):
//   "1" = görüldü (birincil düğme / zil menüsü açıldı) — zildeki kırmızı nokta söner,
//   "m" = modal kapatıldı ("Daha sonra", X, Esc, arka plan) — modal bir daha açılmaz
//         ama zildeki nokta, kullanıcı zili açana dek kalır.
// Aynı anahtarlar bildirim zilindeki "Duyurular" bölümüyle paylaşılır.

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import Icon from "./Icon";
import type { Duyuru } from "./types";

/** Bir duyuru görüldüğünde/kapatıldığında pencereye yayınlanan olay (zil dinler). */
export const DUYURU_EVENT = "olga-duyuru-goruldu";

export function seenKey(id: string, username: string): string {
  return `olga-duyuru:${id}:${username}`;
}

function oku(id: string, username: string): string | null {
  try { return localStorage.getItem(seenKey(id, username)); } catch { return null; }
}

function yayinla(id: string): void {
  try { window.dispatchEvent(new CustomEvent(DUYURU_EVENT, { detail: { id } })); } catch { /* yok */ }
}

/** Tam görüldü mü (zil rozeti için). */
export function isSeen(id: string, username: string): boolean {
  return oku(id, username) === "1";
}

/** Modal olarak bir daha açılmamalı mı (görüldü VEYA "daha sonra" dendi). */
export function isDismissed(id: string, username: string): boolean {
  const v = oku(id, username);
  return v === "1" || v === "m";
}

export function markSeen(id: string, username: string): void {
  try { localStorage.setItem(seenKey(id, username), "1"); } catch { /* depolama yok */ }
  yayinla(id);
}

/** "Daha sonra": modal kapanır, zil noktası kalır. Tam görülmüşü geri almaz. */
export function markDismissed(id: string, username: string): void {
  try {
    const k = seenKey(id, username);
    if (localStorage.getItem(k) !== "1") localStorage.setItem(k, "m");
  } catch { /* depolama yok */ }
  yayinla(id);
}

const GECIKME_MS = 600; // sayfa önce boyansın
const ODAK_SECICI = 'a[href], button:not([disabled]), [tabindex]:not([tabindex="-1"])';

export default function AnnouncementModal({ duyurular, username }: { duyurular: Duyuru[]; username: string }) {
  const [aktif, setAktif] = useState<Duyuru | null>(null);
  const gosterildiRef = useRef(false); // ziyaret başına en fazla bir modal
  const dialogRef = useRef<HTMLDivElement>(null);
  const linkRef = useRef<HTMLAnchorElement>(null);
  const btnRef = useRef<HTMLButtonElement>(null);
  const oncekiOdakRef = useRef<Element | null>(null);

  useEffect(() => {
    if (gosterildiRef.current) return;
    if (!duyurular.some((d) => !isDismissed(d.id, username))) return;
    const t = setTimeout(() => {
      if (gosterildiRef.current) return;
      // Bu arada (ör. zil açılıp) görülmüş olabilir — yeniden bak
      const ilk = duyurular.find((d) => !isDismissed(d.id, username));
      if (!ilk) return;
      gosterildiRef.current = true;
      setAktif(ilk);
    }, GECIKME_MS);
    return () => clearTimeout(t);
  }, [duyurular, username]);

  // tam=true: "Anladım" / birincil bağlantı (görüldü). tam=false: "Daha sonra", X, Esc, arka plan.
  const bitir = useCallback((tam: boolean) => {
    if (aktif) (tam ? markSeen : markDismissed)(aktif.id, username);
    setAktif(null);
  }, [aktif, username]);
  const kapat = useCallback(() => bitir(false), [bitir]);
  const gordum = useCallback(() => bitir(true), [bitir]);

  // Esc kapatır, Tab modal içinde döner, gövde kaydırması kilitlenir,
  // odak birincil düğmeye gider ve kapanınca eski yerine döner
  useEffect(() => {
    if (!aktif) return;
    oncekiOdakRef.current = document.activeElement;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") { e.preventDefault(); kapat(); return; }
      if (e.key !== "Tab") return;
      const dlg = dialogRef.current;
      if (!dlg) return;
      const liste = dlg.querySelectorAll<HTMLElement>(ODAK_SECICI);
      if (liste.length === 0) return;
      const ilk = liste[0];
      const son = liste[liste.length - 1];
      const simdiki = document.activeElement;
      const disarida = !dlg.contains(simdiki);
      if (e.shiftKey && (disarida || simdiki === ilk)) { e.preventDefault(); son.focus(); }
      else if (!e.shiftKey && (disarida || simdiki === son)) { e.preventDefault(); ilk.focus(); }
    };
    window.addEventListener("keydown", onKey);
    const prev = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    (linkRef.current ?? btnRef.current)?.focus();
    return () => {
      window.removeEventListener("keydown", onKey);
      document.body.style.overflow = prev;
      const o = oncekiOdakRef.current;
      if (o instanceof HTMLElement && o !== document.body && document.contains(o)) o.focus();
    };
  }, [aktif, kapat]);

  if (!aktif) return null;

  return (
    <div className="backdrop ann-backdrop no-print" onClick={(e) => { if (e.target === e.currentTarget) kapat(); }}>
      <div
        ref={dialogRef}
        className="ann-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="ann-title"
        aria-describedby="ann-text"
      >
        <button type="button" className="btn icon ghost small ann-close" onClick={kapat} aria-label="Kapat">
          <Icon name="x" size={16} />
        </button>
        <span className="ann-icon"><Icon name={aktif.ikon || "sparkles"} size={26} /></span>
        <span className="ann-kicker">Yeni özellik</span>
        <h2 className="ann-title" id="ann-title">{aktif.baslik}</h2>
        <p className="ann-text" id="ann-text">{aktif.metin}</p>
        <div className="ann-actions">
          {aktif.href ? (
            <Link ref={linkRef} className="btn" href={aktif.href} onClick={gordum}>
              {aktif.hrefLabel || "İncele"} <Icon name="arrow-up-right" size={16} />
            </Link>
          ) : (
            <button ref={btnRef} type="button" className="btn" onClick={gordum}>Anladım</button>
          )}
          <button type="button" className="btn secondary" onClick={kapat}>Daha sonra</button>
        </div>
      </div>
    </div>
  );
}
