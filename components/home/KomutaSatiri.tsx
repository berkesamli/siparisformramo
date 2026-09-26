"use client";

// Ana sayfa komuta satırı: tek kutuya müşteri adı / profil kodu / sipariş no
// yazılır, cevap yerinde gelir (sayfa değişmez).
// - Müşteri: Mikro bakiyesi (BakiyeChip, tembel), son siparişi, telefonu,
//   "bu müşteriye yeni sipariş" düğmesi.
// - Stok: Ankara/İstanbul boy sayısı ve liste fiyatı tek satırda.
// - Sipariş: durum rozeti, tıklayınca detay.
// Veri /api/search'ten gelir (Ctrl+K paletiyle aynı uç). Kutu boşken "son
// baktıkların" çipleri (tarayıcıda saklanır). Yanında büyük "Yeni Sipariş" ve
// "Metin Yapıştır" düğmeleri.

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import Icon, { type IconName } from "@/components/shell/Icon";
import type { SearchHit, SearchKind } from "@/components/shell/SearchPalette";
import { STATUS_LABELS, type OrderStatus } from "@/lib/orders";
import BakiyeChip from "./BakiyeChip";

const SON_KEY = "olga-son-bakilan";
const SON_MAX = 8;
interface SonBakilan { q: string; label: string; kind: SearchKind }

const GRUP: { kind: SearchKind; label: string; icon: IconName }[] = [
  { kind: "customer", label: "Müşteriler", icon: "users" },
  { kind: "stock", label: "Stok", icon: "package" },
  { kind: "order", label: "Toptan Siparişler", icon: "list" },
  { kind: "retail", label: "Perakende Siparişler", icon: "frame" },
  { kind: "catalog", label: "Çerçeve Profilleri", icon: "layers" },
  { kind: "technical", label: "Teknik Malzeme", icon: "box" },
  { kind: "retailCustomer", label: "Perakende Müşteriler", icon: "user" },
];
const IKON: Record<SearchKind, IconName> = {
  page: "arrow-up-right", order: "list", retail: "frame", customer: "users",
  retailCustomer: "user", stock: "package", catalog: "layers", technical: "box",
};
const DURUM_ROZET: Record<string, string> = { olusturuldu: "warn", hazirlaniyor: "info", yarim: "yarim", tamamlandi: "ok", iptal: "err" };

const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });
const gun = (k: string) => { const [, m, d] = k.split("-"); return `${Number(d)}.${m}`; };

function sonOku(): SonBakilan[] {
  try {
    const raw = localStorage.getItem(SON_KEY);
    const v = raw ? JSON.parse(raw) : [];
    return Array.isArray(v) ? v.slice(0, SON_MAX) : [];
  } catch { return []; }
}

export default function KomutaSatiri() {
  const router = useRouter();
  const [q, setQ] = useState("");
  const [hits, setHits] = useState<SearchHit[]>([]);
  const [loading, setLoading] = useState(false);
  const [son, setSon] = useState<SonBakilan[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    setSon(sonOku());
    // Masaüstünde imleç kutuda başlasın; telefonda klavye kendiliğinden açılmasın
    if (window.matchMedia("(min-width: 900px)").matches) inputRef.current?.focus();
  }, []);

  // Uzak arama — 250 ms gecikmeli, önceki istek iptal (paletle aynı düzen)
  useEffect(() => {
    const term = q.trim();
    if (term.length < 2) { setHits([]); setLoading(false); return; }
    setLoading(true);
    const t = setTimeout(async () => {
      abortRef.current?.abort();
      const ac = new AbortController();
      abortRef.current = ac;
      try {
        const r = await fetch(`/api/search?q=${encodeURIComponent(term)}`, { signal: ac.signal });
        const d = await r.json();
        if (!ac.signal.aborted) setHits(d.ok ? (d.hits as SearchHit[]) : []);
      } catch { /* iptal veya ağ hatası */ }
      finally { if (!ac.signal.aborted) setLoading(false); }
    }, 250);
    return () => clearTimeout(t);
  }, [q]);

  // Tıklanan sonucu "son baktıkların" listesine yaz
  const hatirla = useCallback((h: SearchHit) => {
    const label = h.kind === "order" || h.kind === "retail" ? h.title.split(" · ")[0] : h.title;
    const kayit: SonBakilan = { q: label, label, kind: h.kind };
    setSon((prev) => {
      const next = [kayit, ...prev.filter((x) => x.label !== label)].slice(0, SON_MAX);
      try { localStorage.setItem(SON_KEY, JSON.stringify(next)); } catch { /* özel pencere */ }
      return next;
    });
  }, []);

  const gruplar = useMemo(
    () => GRUP.map((g) => ({ ...g, items: hits.filter((h) => h.kind === g.kind) })).filter((g) => g.items.length),
    [hits]
  );
  const term = q.trim();

  function onKey(e: React.KeyboardEvent<HTMLInputElement>) {
    if (e.key === "Escape") { setQ(""); return; }
    if (e.key === "Enter") {
      e.preventDefault();
      const ilk = hits.find((h) => h.kind !== "page");
      if (ilk) { hatirla(ilk); router.push(ilk.href); }
    }
  }

  return (
    <section className="card pad-sm ks" aria-label="Komuta satırı">
      <div className="ks-bar">
        <div className="ks-input">
          <Icon name="search" size={20} />
          <input
            ref={inputRef}
            value={q}
            onChange={(e) => setQ(e.target.value)}
            onKeyDown={onKey}
            placeholder="Müşteri adı, profil kodu veya sipariş no…"
            aria-label="Müşteri, stok veya sipariş ara"
            autoComplete="off"
            spellCheck={false}
            enterKeyHint="search"
          />
          {loading && <span className="ks-loading muted small">Aranıyor…</span>}
          {term ? (
            <button type="button" className="btn icon ghost small" onClick={() => { setQ(""); inputRef.current?.focus(); }} aria-label="Temizle">
              <Icon name="x" size={16} />
            </button>
          ) : (
            <span className="kbd ks-kbd" title="Her sayfada Ctrl/⌘+K ile de açılır">⌘K</span>
          )}
        </div>
        <Link href="/panel" className="btn ks-btn"><Icon name="plus" size={18} /> Yeni Sipariş</Link>
        <Link href="/panel?metin=1" className="btn secondary ks-btn" title="WhatsApp mesajı ya da listeyi yapıştır, satırlara çevrilsin">
          <Icon name="copy" size={18} /> Metin Yapıştır
        </Link>
      </div>

      {!term && son.length > 0 && (
        <div className="ks-son">
          <span className="ks-son-label"><Icon name="clock" size={13} /> Son baktıkların</span>
          {son.map((s) => (
            <button key={`${s.kind}-${s.label}`} type="button" className="chip" onClick={() => { setQ(s.q); inputRef.current?.focus(); }}>
              <Icon name={IKON[s.kind] || "search"} size={13} /> {s.label}
            </button>
          ))}
          <button
            type="button"
            className="btn xs ghost"
            onClick={() => { setSon([]); try { localStorage.removeItem(SON_KEY); } catch { /* yok */ } }}
            aria-label="Son baktıklarını temizle"
          >
            Temizle
          </button>
        </div>
      )}
      {!term && son.length === 0 && (
        <p className="ks-hint">
          Müşteri adı yazınca bakiyesi ve son siparişi, profil kodu yazınca depo stoğu ve liste fiyatı,
          sipariş numarası yazınca durumu burada görünür.
        </p>
      )}

      {term.length >= 2 && (
        <div className="ks-results" aria-live="polite">
          {gruplar.map((g) => (
            <div key={g.kind} className="ks-group">
              <div className="ks-group-title"><Icon name={g.icon} size={13} /> {g.label}</div>
              {g.items.map((h) =>
                h.kind === "customer"
                  ? <MusteriSatiri key={`${h.kind}-${h.href}`} h={h} onGo={hatirla} />
                  : <BasitSatir key={`${h.kind}-${h.href}-${h.title}`} h={h} onGo={hatirla} />
              )}
            </div>
          ))}
          {!loading && gruplar.length === 0 && (
            <div className="ks-empty">
              “{term}” için sonuç yok. Yazımı kontrol edin ya da <Link href="/musteriler">müşteri defterine</Link> bakın.
            </div>
          )}
        </div>
      )}
    </section>
  );
}

function MusteriSatiri({ h, onGo }: { h: SearchHit; onGo: (h: SearchHit) => void }) {
  const id = h.data?.customerId;
  const tel = h.data?.phone;
  const so = h.data?.sonSiparis;
  return (
    <div className="ks-row ks-musteri">
      <span className="avatar ks-avatar" aria-hidden>{(h.title || "?").slice(0, 1).toLocaleUpperCase("tr-TR")}</span>
      <div className="ks-main">
        <Link href={h.href} className="ks-title" onClick={() => onGo(h)}>{h.title}</Link>
        {h.sub && <span className="ks-sub">{h.sub}</span>}
        <div className="ks-facts">
          {id && <BakiyeChip customerId={id} />}
          {so ? (
            <Link href={so.href} className="ks-fact" onClick={() => onGo(h)} title="Son siparişi aç">
              <Icon name="list" size={12} /> Son sipariş {gun(so.dateKey)} · {tl(so.net)}{" "}
              <span className={`badge ${DURUM_ROZET[so.status] || ""}`}>{STATUS_LABELS[so.status as OrderStatus] || so.status}</span>
            </Link>
          ) : (
            <span className="ks-fact muted">Son 6 ayda sipariş yok</span>
          )}
          {h.meta && <span className="badge brand">{h.meta}</span>}
        </div>
      </div>
      <div className="ks-actions">
        {tel && (
          <a href={`tel:${tel.replace(/[^\d+]/g, "")}`} className="btn icon secondary small" title={tel} aria-label={`Ara: ${tel}`}>
            <Icon name="phone" size={15} />
          </a>
        )}
        {id && (
          <Link href={`/panel?musteri=${encodeURIComponent(id)}`} className="btn small" onClick={() => onGo(h)}>
            <Icon name="plus" size={15} /> Yeni Sipariş
          </Link>
        )}
      </div>
    </div>
  );
}

function BasitSatir({ h, onGo }: { h: SearchHit; onGo: (h: SearchHit) => void }) {
  return (
    <Link href={h.href} className="ks-row ks-link" onClick={() => onGo(h)}>
      <span className="ks-ico"><Icon name={IKON[h.kind]} size={17} /></span>
      <span className="ks-main">
        <span className="ks-title">{h.title}</span>
        {h.sub && <span className="ks-sub">{h.sub}</span>}
      </span>
      {h.meta && <span className={`badge ${h.metaKind || ""}`}>{h.meta}</span>}
      <Icon name="chevron-right" size={16} className="ks-chev" />
    </Link>
  );
}
