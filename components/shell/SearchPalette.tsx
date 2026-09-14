"use client";

// Genel arama paleti (Ctrl/⌘+K): sayfalar (istemcide), siparişler, perakende
// siparişler, müşteriler, stok ve katalog (/api/search). Klavye: ↑↓ gez, ↵ aç, Esc kapat.

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import Icon, { type IconName } from "./Icon";
import { flatNav } from "./nav-config";
import type { ShellUser } from "./types";

export type SearchKind = "page" | "order" | "retail" | "customer" | "retailCustomer" | "stock" | "catalog" | "technical";

export interface SearchHit {
  kind: SearchKind;
  title: string;
  sub?: string;
  href: string;
  meta?: string;
  metaKind?: "ok" | "warn" | "err" | "info" | "brand";
}

const KIND_LABEL: Record<SearchKind, string> = {
  page: "Sayfalar",
  order: "Toptan Siparişler",
  retail: "Perakende Siparişler",
  customer: "Müşteriler",
  retailCustomer: "Perakende Müşteriler",
  stock: "Stok",
  catalog: "Çerçeve Profilleri",
  technical: "Teknik Malzeme",
};
const KIND_ICON: Record<SearchKind, IconName> = {
  page: "arrow-up-right",
  order: "list",
  retail: "frame",
  customer: "users",
  retailCustomer: "user",
  stock: "package",
  catalog: "layers",
  technical: "box",
};
const KIND_ORDER: SearchKind[] = ["page", "order", "retail", "customer", "retailCustomer", "stock", "catalog", "technical"];
const RECENT_KEY = "olga-recent-search";

function norm(s: string): string {
  return String(s || "")
    .replace(/[çÇ]/g, "c").replace(/[ğĞ]/g, "g").replace(/[ıİ]/g, "i")
    .replace(/[öÖ]/g, "o").replace(/[şŞ]/g, "s").replace(/[üÜ]/g, "u")
    .toLowerCase().trim();
}

export default function SearchPalette({ user, onClose }: { user: ShellUser; onClose: () => void }) {
  const router = useRouter();
  const [q, setQ] = useState("");
  const [remote, setRemote] = useState<SearchHit[]>([]);
  const [loading, setLoading] = useState(false);
  const [idx, setIdx] = useState(0);
  const [recent, setRecent] = useState<string[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    inputRef.current?.focus();
    try {
      const raw = localStorage.getItem(RECENT_KEY);
      if (raw) setRecent(JSON.parse(raw));
    } catch { /* yok */ }
    const prev = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = prev; };
  }, []);

  // Sayfa eşleşmeleri — istemcide, anında
  const pages = useMemo<SearchHit[]>(() => {
    const nq = norm(q);
    const all = flatNav(user);
    const list = nq
      ? all.filter((it) => norm(it.label).includes(nq) || (it.keywords || []).some((k) => norm(k).includes(nq)))
      : all.slice(0, 6);
    return list.slice(0, 6).map((it) => ({ kind: "page", title: it.label, sub: it.href, href: it.href }));
  }, [q, user]);

  // Uzak arama — 250 ms gecikmeli, önceki istek iptal
  useEffect(() => {
    const term = q.trim();
    if (term.length < 2) { setRemote([]); setLoading(false); return; }
    setLoading(true);
    const t = setTimeout(async () => {
      abortRef.current?.abort();
      const ac = new AbortController();
      abortRef.current = ac;
      try {
        const r = await fetch(`/api/search?q=${encodeURIComponent(term)}`, { signal: ac.signal });
        const d = await r.json();
        if (!ac.signal.aborted) setRemote(d.ok ? (d.hits as SearchHit[]) : []);
      } catch { /* iptal veya ağ hatası */ }
      finally { if (!ac.signal.aborted) setLoading(false); }
    }, 250);
    return () => clearTimeout(t);
  }, [q]);

  const hits = useMemo(() => {
    const all = [...pages, ...remote];
    all.sort((a, b) => KIND_ORDER.indexOf(a.kind) - KIND_ORDER.indexOf(b.kind));
    return all;
  }, [pages, remote]);

  useEffect(() => { setIdx(0); }, [hits.length, q]);

  const go = useCallback((h: SearchHit) => {
    const term = q.trim();
    if (term) {
      try {
        const next = [term, ...recent.filter((r) => r !== term)].slice(0, 6);
        localStorage.setItem(RECENT_KEY, JSON.stringify(next));
      } catch { /* yok */ }
    }
    onClose();
    router.push(h.href);
  }, [q, recent, onClose, router]);

  function onKey(e: React.KeyboardEvent) {
    if (e.key === "ArrowDown") { e.preventDefault(); setIdx((i) => Math.min(hits.length - 1, i + 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setIdx((i) => Math.max(0, i - 1)); }
    else if (e.key === "Enter") { e.preventDefault(); if (hits[idx]) go(hits[idx]); }
    else if (e.key === "Escape") { e.preventDefault(); onClose(); }
  }

  // Seçili öğe görünür kalsın
  useEffect(() => {
    const el = document.querySelector<HTMLElement>(`[data-sp-idx="${idx}"]`);
    el?.scrollIntoView({ block: "nearest" });
  }, [idx]);

  let flat = -1;
  const grouped = KIND_ORDER.map((k) => ({ kind: k, items: hits.filter((h) => h.kind === k) })).filter((g) => g.items.length);

  return (
    <div className="sp-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="sp" role="dialog" aria-modal="true" aria-label="Genel arama">
        <div className="sp-input">
          <Icon name="search" size={20} />
          <input
            ref={inputRef}
            value={q}
            onChange={(e) => setQ(e.target.value)}
            onKeyDown={onKey}
            placeholder={user.role === "staff" ? "Sipariş no, müşteri, profil kodu, sayfa…" : "Profil kodu, ürün, sayfa…"}
            aria-label="Arama"
            autoComplete="off"
            spellCheck={false}
          />
          {loading && <span className="muted small">Aranıyor…</span>}
          <button type="button" className="btn icon ghost small" onClick={onClose} aria-label="Kapat"><Icon name="x" size={18} /></button>
        </div>
        <div className="sp-body">
          {!q.trim() && recent.length > 0 && (
            <>
              <div className="sp-group-title">Son Aramalar</div>
              <div className="row" style={{ padding: "2px 10px 8px" }}>
                {recent.map((r) => (
                  <button key={r} type="button" className="chip" onClick={() => setQ(r)}>
                    <Icon name="clock" size={13} /> {r}
                  </button>
                ))}
              </div>
            </>
          )}
          {grouped.map((g) => (
            <div key={g.kind}>
              <div className="sp-group-title">{KIND_LABEL[g.kind]}</div>
              {g.items.map((h) => {
                flat += 1;
                const i = flat;
                return (
                  <a
                    key={`${h.kind}-${h.href}-${h.title}`}
                    href={h.href}
                    data-sp-idx={i}
                    className={`sp-item ${i === idx ? "active" : ""}`}
                    onMouseEnter={() => setIdx(i)}
                    onClick={(e) => { e.preventDefault(); go(h); }}
                  >
                    <span className="sp-item-icon"><Icon name={KIND_ICON[h.kind]} size={17} /></span>
                    <span className="sp-item-main">
                      <span className="sp-item-title">{h.title}</span>
                      {h.sub && <span className="sp-item-sub">{h.sub}</span>}
                    </span>
                    {h.meta && <span className={`badge ${h.metaKind || ""} sp-item-meta`}>{h.meta}</span>}
                  </a>
                );
              })}
            </div>
          ))}
          {q.trim().length >= 2 && !loading && hits.length === 0 && (
            <div className="sp-empty">“{q}” için sonuç bulunamadı.</div>
          )}
          {q.trim().length < 2 && (
            <div className="sp-empty" style={{ paddingTop: 12 }}>
              {user.role === "staff"
                ? "Sipariş numarası (OLG-2026-042), müşteri adı, telefon, profil kodu (GC065) veya sayfa adı yazın."
                : "Profil kodu (GC065-1473), ürün adı veya sayfa adı yazın."}
            </div>
          )}
        </div>
        <div className="sp-foot">
          <span><span className="kbd">↑</span> <span className="kbd">↓</span> gezin</span>
          <span><span className="kbd">↵</span> aç</span>
          <span><span className="kbd">esc</span> kapat</span>
          <span className="spacer" />
          <span>Olga Çerçeve</span>
        </div>
      </div>
    </div>
  );
}
