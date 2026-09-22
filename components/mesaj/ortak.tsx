"use client";

import Icon from "@/components/shell/Icon";
import type { Kanal } from "@/lib/mesaj/tur";

/** Kanal ikonu (WhatsApp yeşil, Instagram mor, e-posta mavi — renk CSS'te). */
export function KanalIkon({ kanal, size = 16 }: { kanal: Kanal; size?: number }) {
  if (kanal === "whatsapp") {
    return (
      <svg width={size} height={size} viewBox="0 0 24 24" fill="currentColor" aria-hidden>
        <path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5.1-1.3A10 10 0 1 0 12 2zm0 1.8a8.2 8.2 0 0 1 0 16.4c-1.5 0-2.9-.4-4.1-1.1l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 0 1 12 3.8zm-3 4.4c-.2 0-.5 0-.7.3-.3.3-1 1-1 2.4s1 2.8 1.2 3c.1.2 2 3.1 4.9 4.3 2.4 1 2.9.8 3.4.7.5 0 1.7-.7 1.9-1.4.2-.7.2-1.2.2-1.4-.1-.1-.3-.2-.6-.3l-2-.9c-.3-.1-.5-.2-.7.1l-.9 1.1c-.2.2-.3.2-.6.1a6.7 6.7 0 0 1-3.3-2.9c-.3-.4.3-.4.8-1.5.1-.2 0-.4 0-.5l-.9-2.2c-.2-.6-.5-.5-.7-.5z" />
      </svg>
    );
  }
  if (kanal === "instagram") {
    return (
      <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
        <rect x="3" y="3" width="18" height="18" rx="5" /><circle cx="12" cy="12" r="4" /><circle cx="17.5" cy="6.5" r="1" fill="currentColor" stroke="none" />
      </svg>
    );
  }
  return <Icon name="mail" size={size} />;
}

const AY = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];

/** Liste için kısa zaman: 14:05 · Dün · 12 Eyl · 3 Oca 25 */
export function zamanKisa(iso: string): string {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const simdi = new Date();
  const ayni = d.toDateString() === simdi.toDateString();
  if (ayni) return d.toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit" });
  const dun = new Date(simdi); dun.setDate(simdi.getDate() - 1);
  if (d.toDateString() === dun.toDateString()) return "Dün";
  if (d.getFullYear() === simdi.getFullYear()) return `${d.getDate()} ${AY[d.getMonth()]}`;
  return `${d.getDate()} ${AY[d.getMonth()]} ${String(d.getFullYear()).slice(-2)}`;
}

export function zamanTam(iso: string): string {
  const d = new Date(iso);
  return isNaN(d.getTime()) ? "" : d.toLocaleString("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
}

export function gunBasligi(iso: string): string {
  const d = new Date(iso);
  const simdi = new Date();
  if (d.toDateString() === simdi.toDateString()) return "Bugün";
  const dun = new Date(simdi); dun.setDate(simdi.getDate() - 1);
  if (d.toDateString() === dun.toDateString()) return "Dün";
  return d.toLocaleDateString("tr-TR", { day: "numeric", month: "long", year: d.getFullYear() === simdi.getFullYear() ? undefined : "numeric" });
}
