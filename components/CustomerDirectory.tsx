"use client";

// Müşteriler sayfası: arama, şehir/şube filtreleri, özet sayaçlar ve
// tıklanınca müşteri kartına giden satırlar. Masaüstünde sütunlu satır,
// telefonda kart düzeni (customers.css). Ekle/düzenle CustomerForm modalı.

import { useCallback, useEffect, useMemo, useState } from "react";
import { useRouter } from "next/navigation";
import Icon from "@/components/shell/Icon";
import CustomerForm from "@/components/CustomerForm";
import { customerTitle, normalizeCity, musteriBolgesi, bolgeler as varsayilanBolgeler, BOLGE_SIRASI, type Bolge, type BolgeInfo, type Customer } from "@/lib/customers";
import { eslesir } from "@/lib/search-norm";
import { initials } from "@/components/shell/nav-config";

const SAYFA = 40;

export default function CustomerDirectory({ initialBolge = "" }: { initialBolge?: "" | Bolge }) {
  const router = useRouter();
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [loading, setLoading] = useState(true);
  const [blobOk, setBlobOk] = useState(true);
  const [query, setQuery] = useState("");
  const [city, setCity] = useState("");
  const [bolge, setBolge] = useState<"" | Bolge>(initialBolge);
  const [bolgeTanim, setBolgeTanim] = useState<Record<Bolge, BolgeInfo>>(() => varsayilanBolgeler());
  const [mikro, setMikro] = useState<"" | "bagli" | "bagsiz">("");
  const [sort, setSort] = useState<"ad" | "yeni">("ad");
  const [limit, setLimit] = useState(SAYFA);
  const [form, setForm] = useState<null | "yeni" | Customer>(null);
  const [msg, setMsg] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/musteriler");
      const d = await res.json();
      if (res.ok) {
        setCustomers(d.customers || []);
        setBlobOk(d.blob !== false);
        if (d.bolgeler) setBolgeTanim(d.bolgeler);
      }
    } finally {
      setLoading(false);
    }
  }, []);
  useEffect(() => { load(); }, [load]);

  // Şehirler (kayıtlardan) ve sayaçlar
  const cities = useMemo(() => {
    const m = new Map<string, { label: string; n: number }>();
    customers.forEach((c) => {
      if (!c.city) return;
      const k = normalizeCity(c.city);
      const cur = m.get(k);
      m.set(k, { label: cur?.label || c.city, n: (cur?.n || 0) + 1 });
    });
    return [...m.entries()].sort((a, b) => b[1].n - a[1].n || a[1].label.localeCompare(b[1].label, "tr"));
  }, [customers]);
  const sayac = useMemo(() => {
    const b: Record<Bolge, number> = { ankara: 0, istanbul: 0, tasra: 0 };
    customers.forEach((c) => { b[musteriBolgesi(c)]++; });
    return { toplam: customers.length, bolge: b, mikro: customers.filter((c) => c.mikroCariKod).length };
  }, [customers]);

  const filtered = useMemo(() => {
    const q = query.trim();
    const list = customers.filter((c) => {
      if (city && normalizeCity(c.city) !== city) return false;
      if (bolge && musteriBolgesi(c) !== bolge) return false;
      if (mikro === "bagli" && !c.mikroCariKod) return false;
      if (mikro === "bagsiz" && c.mikroCariKod) return false;
      if (!q) return true;
      return eslesir(q, customerTitle(c), c.company, c.firstName, c.lastName, c.email, c.phone, c.addr1, c.addr2, c.city, c.district, c.mikroUnvan, c.mikroCariKod);
    });
    if (sort === "yeni") list.sort((a, b) => (b.updatedAt || "").localeCompare(a.updatedAt || ""));
    return list;
  }, [customers, query, city, bolge, mikro, sort]);

  useEffect(() => { setLimit(SAYFA); }, [query, city, bolge, mikro, sort]);

  const filtreVar = Boolean(query || city || bolge || mikro);
  const kartAc = (c: Customer) => router.push(`/musteriler/kart?id=${encodeURIComponent(c.id)}`);

  function onSaved(c: Customer) {
    setForm(null);
    setMsg(customers.some((x) => x.id === c.id) ? "Müşteri güncellendi." : "Müşteri eklendi.");
    setCustomers((list) => {
      const i = list.findIndex((x) => x.id === c.id);
      const next = i >= 0 ? list.map((x) => (x.id === c.id ? c : x)) : [...list, c];
      return next.sort((a, b) => customerTitle(a).localeCompare(customerTitle(b), "tr", { sensitivity: "base" }));
    });
    setTimeout(() => setMsg(""), 3000);
  }

  return (
    <div>
      {/* Araç çubuğu */}
      <div className="cd-toolbar">
        <div className="cd-search">
          <Icon name="search" size={18} />
          <input
            placeholder="Firma, kişi, telefon, şehir ya da Mikro cari kodu…"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            autoComplete="off"
            aria-label="Müşteri ara"
          />
          {query && (
            <button type="button" className="cd-search-clear" onClick={() => setQuery("")} aria-label="Aramayı temizle">
              <Icon name="x" size={16} />
            </button>
          )}
        </div>
        <div className="seg" aria-label="Sıralama">
          <button type="button" className={sort === "ad" ? "active" : ""} onClick={() => setSort("ad")}>A → Z</button>
          <button type="button" className={sort === "yeni" ? "active" : ""} onClick={() => setSort("yeni")}>Son güncellenen</button>
        </div>
        <button type="button" className="btn" onClick={() => setForm("yeni")}>
          <Icon name="plus" size={16} /> Yeni Müşteri
        </button>
      </div>

      {/* Bölge sayaçları — tıklayınca filtre; altında bölgeyle ilgilenen satışçı */}
      <div className="cd-stats">
        <button type="button" className={`cd-stat tum ${!bolge && !mikro ? "sel" : ""}`} onClick={() => { setBolge(""); setMikro(""); }}>
          <b>{sayac.toplam}</b><span>Tüm müşteriler</span>
        </button>
        {BOLGE_SIRASI.map((b) => (
          <button key={b} type="button" className={`cd-stat ${bolge === b ? "sel" : ""}`} onClick={() => setBolge(bolge === b ? "" : b)} title={`${bolgeTanim[b].label} müşterileri`}>
            <b>{sayac.bolge[b]}</b><span>{bolgeTanim[b].label} müşterileri</span>
          </button>
        ))}
        <button type="button" className={`cd-stat ${mikro === "bagli" ? "sel" : ""}`} onClick={() => setMikro(mikro === "bagli" ? "" : "bagli")} title="Mikro cari kartıyla eşleştirilmiş müşteriler">
          <b>{sayac.mikro}</b><span>Mikro&apos;ya bağlı</span>
        </button>
      </div>

      {/* Şehir çipleri */}
      {cities.length > 1 && (
        <div className="cd-chips">
          <button type="button" className={`chip ${city === "" ? "sel" : ""}`} onClick={() => setCity("")}>Tüm şehirler</button>
          {cities.slice(0, 12).map(([k, v]) => (
            <button key={k} type="button" className={`chip ${city === k ? "sel" : ""}`} onClick={() => setCity(city === k ? "" : k)}>
              {v.label} <b>{v.n}</b>
            </button>
          ))}
          {mikro !== "" && (
            <button type="button" className="chip sel" onClick={() => setMikro("")}>
              {mikro === "bagli" ? "Mikro'ya bağlı" : "Mikro'ya bağlı değil"} <Icon name="x" size={12} />
            </button>
          )}
        </div>
      )}

      {!blobOk && <div className="notice info">Kalıcı depolama yapılandırılmadığı için müşteriler kaydedilemiyor.</div>}
      {msg && <div className="notice ok">{msg}</div>}

      {/* Liste */}
      {loading ? (
        <div className="cd-list" aria-busy="true">
          {[0, 1, 2, 3, 4].map((i) => <div key={i} className="skeleton" style={{ height: 68, borderRadius: 11 }} />)}
        </div>
      ) : filtered.length === 0 ? (
        <div className="card">
          <div className="empty">
            <div className="empty-icon"><Icon name="users" size={22} /></div>
            <strong>{customers.length === 0 ? "Henüz müşteri yok" : "Bu filtreye uyan müşteri yok"}</strong>
            {customers.length === 0 ? "“Yeni Müşteri” ile ilk kaydı ekleyin." : "Arama kelimesini kısaltın ya da filtreleri kaldırın."}
            {filtreVar && (
              <div style={{ marginTop: 12 }}>
                <button type="button" className="btn secondary small" onClick={() => { setQuery(""); setCity(""); setBolge(""); setMikro(""); }}>Filtreleri temizle</button>
              </div>
            )}
          </div>
        </div>
      ) : (
        <>
          <div className="muted" style={{ fontSize: 12.5, marginBottom: 8 }}>
            {filtered.length} müşteri{filtreVar ? ` (toplam ${customers.length})` : ""}
          </div>
          <div className="cd-list">
            {filtered.slice(0, limit).map((c) => {
              const ad = customerTitle(c);
              const kisi = `${c.firstName || ""} ${c.lastName || ""}`.trim();
              const konum = [c.district, c.city].filter(Boolean).join(" / ");
              const b = musteriBolgesi(c);
              return (
                <div
                  key={c.id}
                  className="cd-row"
                  role="link"
                  tabIndex={0}
                  onClick={() => kartAc(c)}
                  onKeyDown={(e) => { if (e.key === "Enter") kartAc(c); }}
                >
                  <span className={`cd-avatar ${b}`} aria-hidden>{initials(c.company || kisi || ad)}</span>
                  <div className="cd-main">
                    <div className="cd-name">{c.company || kisi || ad}</div>
                    <div className="cd-sub">
                      {[c.company && kisi ? kisi : "", konum].filter(Boolean).join(" · ")}
                      {c.phone && <span className="cd-sub-tel">{(c.company && kisi) || konum ? " · " : ""}{c.phone}</span>}
                      {!kisi && !konum && !c.phone && "İletişim bilgisi girilmemiş"}
                    </div>
                  </div>
                  <div className="cd-cell">
                    {c.phone || <span className="muted">Telefon yok</span>}
                    <small>{c.email || konum || " "}</small>
                  </div>
                  <div className="cd-tags">
                    <span className={`bolge ${b}`} title={`${bolgeTanim[b].label} müşterisi`}>{bolgeTanim[b].kisa}</span>
                    {c.mikroCariKod && <span className="badge ok" title={`Mikro: ${c.mikroUnvan || c.mikroCariKod}`}><Icon name="briefcase" size={12} /> Mikro</span>}
                    {c.iskontoPct ? <span className="badge brand">%{c.iskontoPct} isk.</span> : null}
                  </div>
                  <div className="cd-actions" onClick={(e) => e.stopPropagation()}>
                    <a className="btn small secondary icon" href={`/musteriler/kart?id=${encodeURIComponent(c.id)}`} title="Müşteri kartı" aria-label="Müşteri kartı">
                      <Icon name="credit-card" size={15} /><span className="btn-lbl">Kart</span>
                    </a>
                    <a className="btn small secondary icon" href={`/etiket?id=${encodeURIComponent(c.id)}`} title="Kargo etiketi" aria-label="Kargo etiketi">
                      <Icon name="tag" size={15} /><span className="btn-lbl">Etiket</span>
                    </a>
                    <button type="button" className="btn small secondary icon" title="Düzenle" aria-label="Düzenle" onClick={() => setForm(c)}>
                      <Icon name="edit" size={15} /><span className="btn-lbl">Düzenle</span>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
          {filtered.length > limit && (
            <div className="cd-more">
              <button type="button" className="btn secondary" onClick={() => setLimit((n) => n + SAYFA)}>
                Daha fazla göster ({filtered.length - limit} kaldı)
              </button>
            </div>
          )}
        </>
      )}

      {form && (
        <CustomerForm
          initial={form === "yeni" ? null : form}
          onSaved={onSaved}
          onClose={() => setForm(null)}
        />
      )}
    </div>
  );
}
