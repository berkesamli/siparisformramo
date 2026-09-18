"use client";

// Kargo etiketi (150×100 mm): kayıtlı müşteriyi seç, gönderici şubeyi
// belirle, PDF yazdır. Müşteri ekleme/düzenleme ana yeri /musteriler; burada
// hızlı düzenleme (CustomerForm modalı) ve yeni kayıt kısayolu var.
// ?id= ile gelindiğinde (müşteri kartı / dizin) o müşteri seçili açılır.

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import CustomerForm from "@/components/CustomerForm";
import {
  branchInfo,
  customerTitle,
  normalizeCity,
  type Branch,
  type Customer,
} from "@/lib/customers";
import { eslesir } from "@/lib/search-norm";

export default function LabelManager({ preselectId = "" }: { preselectId?: string }) {
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [loading, setLoading] = useState(true);
  const [blobOk, setBlobOk] = useState(true);
  const [query, setQuery] = useState("");
  const [cityFilter, setCityFilter] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(preselectId || null);
  const [form, setForm] = useState<null | "yeni" | Customer>(null);
  const [msg, setMsg] = useState("");
  const [copies, setCopies] = useState("1");
  const layoutRef = useRef<HTMLDivElement | null>(null);
  const previewRef = useRef<HTMLElement | null>(null);

  // Dar ekranda önizleme listenin altında kalır — müşteri seçilince oraya kaydır.
  function selectCustomer(id: string) {
    setSelectedId(id);
    if (typeof window === "undefined") return;
    const layout = layoutRef.current;
    const cols = layout ? window.getComputedStyle(layout).gridTemplateColumns.trim().split(/\s+/).length : 0;
    const stacked = cols ? cols === 1 : window.matchMedia("(max-width: 1080px)").matches;
    if (stacked) previewRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  const load = useCallback(async (preferId?: string) => {
    setLoading(true);
    try {
      const res = await fetch("/api/musteriler");
      const d = await res.json();
      if (res.ok) {
        const list: Customer[] = d.customers || [];
        setCustomers(list);
        setBlobOk(d.blob !== false);
        setSelectedId((cur) => {
          const want = preferId || cur;
          if (want && list.some((c) => c.id === want)) return want;
          return list.length ? list[0].id : null;
        });
      }
    } finally {
      setLoading(false);
    }
  }, []);
  useEffect(() => { load(); }, [load]);

  const cities = useMemo(() => {
    const map = new Map<string, { label: string; n: number }>();
    customers.forEach((c) => {
      if (!c.city) return;
      const k = normalizeCity(c.city);
      const cur = map.get(k);
      map.set(k, { label: cur?.label || c.city, n: (cur?.n || 0) + 1 });
    });
    return [...map.entries()].sort((a, b) => b[1].n - a[1].n || a[1].label.localeCompare(b[1].label, "tr"));
  }, [customers]);

  const filtered = useMemo(() => {
    const q = query.trim();
    return customers.filter((c) => {
      if (cityFilter && normalizeCity(c.city) !== cityFilter) return false;
      if (!q) return true;
      return eslesir(q, customerTitle(c), c.company, c.firstName, c.lastName, c.email, c.phone, c.addr1, c.addr2, c.city, c.district);
    });
  }, [customers, query, cityFilter]);

  const selected = customers.find((c) => c.id === selectedId) || null;

  function onSaved(c: Customer) {
    setForm(null);
    setMsg(customers.some((x) => x.id === c.id) ? "Müşteri güncellendi." : "Müşteri eklendi.");
    setCustomers((list) => {
      const i = list.findIndex((x) => x.id === c.id);
      const next = i >= 0 ? list.map((x) => (x.id === c.id ? c : x)) : [...list, c];
      return next.sort((a, b) => customerTitle(a).localeCompare(customerTitle(b), "tr", { sensitivity: "base" }));
    });
    setSelectedId(c.id);
    setTimeout(() => setMsg(""), 3000);
  }

  async function subeDegistir(b: Branch) {
    if (!selected || selected.branch === b) return;
    const upd = { ...selected, branch: b };
    setCustomers((list) => list.map((x) => (x.id === upd.id ? upd : x)));
    await fetch("/api/musteriler", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(upd),
    });
  }

  return (
    <div className="lbl-layout" ref={layoutRef}>
      {/* ---- Sol: müşteri seçimi ---- */}
      <div style={{ minWidth: 0 }}>
        <div className="card pad-sm">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="users" size={16} /></span>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <h2>Alıcı Seç</h2>
              <span className="lbl-count">{filtered.length}</span>
            </div>
            <span className="spacer" />
            <div className="card-head-actions">
              <button type="button" className="btn small secondary" onClick={() => setForm("yeni")}>
                <Icon name="plus" size={14} /> Yeni
              </button>
            </div>
          </div>

          <div className="cd-search">
            <Icon name="search" size={17} />
            <input
              placeholder="Firma, kişi, telefon ya da adres…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              autoComplete="off"
              aria-label="Müşteri ara"
            />
            {query && (
              <button type="button" className="cd-search-clear" onClick={() => setQuery("")} aria-label="Aramayı temizle"><Icon name="x" size={16} /></button>
            )}
          </div>

          {cities.length > 1 && (
            <div className="lbl-cities">
              <button type="button" className={`lbl-city ${cityFilter === "" ? "sel" : ""}`} onClick={() => setCityFilter("")}>
                Tümü <b>{customers.length}</b>
              </button>
              {cities.slice(0, 10).map(([key, v]) => (
                <button key={key} type="button" className={`lbl-city ${cityFilter === key ? "sel" : ""}`} onClick={() => setCityFilter(cityFilter === key ? "" : key)}>
                  {v.label} <b>{v.n}</b>
                </button>
              ))}
            </div>
          )}

          {!blobOk && <div className="notice info" style={{ marginTop: 10 }}>Kalıcı depolama yapılandırılmadığı için müşteriler kaydedilemiyor.</div>}
          {msg && <div className="notice ok" style={{ marginTop: 10 }}>{msg}</div>}

          <div className="lbl-list">
            {loading ? (
              <div className="empty" style={{ padding: "22px 12px" }}>Yükleniyor...</div>
            ) : filtered.length === 0 ? (
              <div className="empty" style={{ padding: "22px 12px" }}>
                {customers.length === 0 ? "Henüz müşteri yok. “Yeni” ile ekleyin." : "Bu aramaya uyan müşteri yok."}
              </div>
            ) : (
              filtered.map((c) => (
                <div
                  key={c.id}
                  className={`lbl-item ${selectedId === c.id ? "sel" : ""}`}
                  role="button"
                  tabIndex={0}
                  onClick={() => selectCustomer(c.id)}
                  onKeyDown={(e) => { if (e.key === "Enter") selectCustomer(c.id); }}
                >
                  <div className="lbl-item-info">
                    <div className="lbl-item-title">{customerTitle(c)}</div>
                    <div className="lbl-item-meta">
                      {[c.phone, [c.district, c.city].filter(Boolean).join(" / ")].filter(Boolean).join(" · ") || "—"}
                    </div>
                  </div>
                  <span className="lbl-item-actions">
                    <span className={`lbl-branch ${c.branch}`}>{c.branch === "istanbul" ? "İST" : "ANK"}</span>
                    {selectedId === c.id && <Icon name="check-circle" size={16} />}
                  </span>
                </div>
              ))
            )}
          </div>
          <p className="muted" style={{ fontSize: 12, marginTop: 10 }}>
            Müşteri bilgilerini toplu görmek ve kartlarını açmak için <Link href="/musteriler">Müşteriler</Link> sayfasını kullanın.
          </p>
        </div>
      </div>

      {/* ---- Sağ: etiket önizleme ve yazdırma ---- */}
      <aside ref={previewRef} style={{ minWidth: 0, scrollMarginTop: "calc(var(--topbar-h) + 12px)" }}>
        <div className="card pad-sm" style={{ position: "sticky", top: "calc(var(--topbar-h) + 16px)" }}>
          <div className="card-head">
            <span className="card-head-icon"><Icon name="tag" size={16} /></span>
            <div>
              <h2>Kargo Etiketi</h2>
              <span className="card-head-sub">150 × 100 mm</span>
            </div>
            {selected && (
              <div className="card-head-actions">
                <button type="button" className="btn small secondary" onClick={() => setForm(selected)} title="Adres ve iletişim bilgilerini düzenle">
                  <Icon name="edit" size={14} /> Düzenle
                </button>
              </div>
            )}
          </div>

          {selected ? (
            <>
              <LabelPreview c={selected} />

              <div style={{ marginTop: 12 }}>
                <label>Gönderici Şube</label>
                <div className="seg" style={{ width: "100%", display: "flex" }}>
                  {(["ankara", "istanbul"] as Branch[]).map((b) => (
                    <button key={b} type="button" className={selected.branch === b ? "active" : ""} style={{ flex: 1 }} onClick={() => subeDegistir(b)}>
                      {b === "ankara" ? "Ankara" : "İstanbul"}
                    </button>
                  ))}
                </div>
                <p className="muted" style={{ fontSize: 11.5, marginTop: 6 }}>
                  Etiketin üstündeki gönderici adresi seçilen şubeye göre değişir; seçim müşteri kartına kaydedilir.
                </p>
              </div>

              <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "flex-end", minWidth: 0 }}>
                <div style={{ flex: "0 0 90px", maxWidth: "100%" }}>
                  <label>Adet</label>
                  <input type="number" inputMode="numeric" min="1" max="50" value={copies} onChange={(e) => setCopies(e.target.value)} />
                </div>
                <a
                  className="btn"
                  style={{ flex: 1, minWidth: 0 }}
                  href={`/api/etiket/pdf?id=${encodeURIComponent(selected.id)}&adet=${Math.max(1, parseInt(copies) || 1)}`}
                >
                  <Icon name="printer" size={16} /> Etiket PDF
                </a>
              </div>
              <div className="row" style={{ marginTop: 10, gap: 6 }}>
                <Link href={`/musteriler/kart?id=${encodeURIComponent(selected.id)}`} className="btn small ghost">
                  <Icon name="credit-card" size={14} /> Müşteri kartı
                </Link>
              </div>
            </>
          ) : (
            <div className="empty" style={{ padding: "22px 12px" }}>
              <div className="empty-icon"><Icon name="tag" size={22} /></div>
              Etiketi görmek için listeden bir müşteri seçin.
            </div>
          )}
        </div>
      </aside>

      {form && (
        <CustomerForm initial={form === "yeni" ? null : form} onSaved={onSaved} onClose={() => setForm(null)} />
      )}
    </div>
  );
}

/* ---- Ekrandaki etiket önizlemesi (PDF ile aynı düzen) ----
   Fiziksel kâğıt etiketi simüle eder: .lbl-preview renkleri (beyaz zemin,
   siyah yazı) labels.css'te bilerek sabittir, temadan etkilenmez. */
function LabelPreview({ c }: { c: Customer }) {
  const b = branchInfo(c.branch);
  const lines = [
    c.addr1,
    c.addr2,
    [c.district, c.city].filter(Boolean).join(" / "),
    [c.postalCode, c.country].filter(Boolean).join(" "),
  ].filter(Boolean);

  return (
    <div className="lbl-preview" style={{ maxWidth: "100%", overflow: "hidden" }}>
      <div className="lbl-hdr">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/logo.png" alt="Olga Çerçeve" />
        <div style={{ minWidth: 0, overflowWrap: "anywhere" }}>
          <div className="lbl-firm">{b.name}</div>
          <div className="lbl-tel">{b.cityTel}</div>
          <div className="lbl-addr">{b.addr1}, {b.addr2}</div>
          <div className="lbl-addr">{b.website}</div>
        </div>
      </div>
      <div className="lbl-to" style={{ minHeight: 0, overflow: "hidden", overflowWrap: "anywhere" }}>
        <div className="lbl-to-title">Alıcı</div>
        <div className="lbl-to-name">{customerTitle(c)}</div>
        {(c.phone || c.email) && (
          <div className="lbl-to-contact">
            {[c.phone && `Tel: ${c.phone}`, c.email && `E-posta: ${c.email}`].filter(Boolean).join("  •  ")}
          </div>
        )}
        <div className="lbl-to-addr">
          {lines.length ? lines.map((l, i) => <div key={i}>{l}</div>) : "—"}
        </div>
        {c.note && <div className="lbl-to-note">{c.note}</div>}
      </div>
    </div>
  );
}
