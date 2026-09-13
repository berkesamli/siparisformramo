"use client";

// Müşteri defteri + 150×100 mm kargo etiketi.
// "etiket" uygulamasının portu: kayıtlı müşteriler, arama, şehir filtresi,
// müşteri ekle/düzenle/sil, gönderici şubesine göre değişen etiket önizlemesi.

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Icon from "@/components/shell/Icon";
import {
  BRANCHES,
  branchInfo,
  customerTitle,
  normalizeCity,
  type Branch,
  type Customer,
} from "@/lib/customers";
import { eslesir } from "@/lib/search-norm";

const BOS: Omit<Customer, "id" | "createdAt" | "updatedAt" | "iskontoPct"> & {
  // Formda metin olarak tutulur ("10" / "12,5"); sunucu sayıya çevirir.
  iskontoPct: string | number;
} = {
  firstName: "",
  lastName: "",
  company: "",
  email: "",
  phone: "",
  addr1: "",
  addr2: "",
  city: "",
  district: "",
  postalCode: "",
  country: "Türkiye",
  branch: "ankara",
  iskontoPct: "",
  note: "",
};

type Form = typeof BOS & { id?: string };

export default function LabelManager() {
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [loading, setLoading] = useState(true);
  const [blobOk, setBlobOk] = useState(true);
  const [query, setQuery] = useState("");
  const [cityFilter, setCityFilter] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [form, setForm] = useState<Form>({ ...BOS });
  const [editing, setEditing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState("");
  const [err, setErr] = useState("");
  const [copies, setCopies] = useState("1");
  const layoutRef = useRef<HTMLDivElement | null>(null);
  const previewRef = useRef<HTMLElement | null>(null);

  // Dar ekranda önizleme listenin altında kalır — müşteri seçilince
  // etiket görünsün diye oraya kaydır (masaüstünde zaten yan yana).
  // Tek sütuna düşüp düşmediği CSS'teki kırılım sayısına bağlı kalmadan,
  // ızgaranın çözümlenmiş sütun sayısından okunur.
  function selectCustomer(id: string) {
    setSelectedId(id);
    if (typeof window === "undefined") return;
    const layout = layoutRef.current;
    const cols = layout
      ? window.getComputedStyle(layout).gridTemplateColumns.trim().split(/\s+/).length
      : 0;
    const stacked = cols ? cols === 1 : window.matchMedia("(max-width: 1080px)").matches;
    if (stacked) {
      previewRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
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

  useEffect(() => {
    load();
  }, [load]);

  // Şehir listesi (kayıtlardan üretilir)
  const cities = useMemo(() => {
    const map = new Map<string, string>();
    customers.forEach((c) => {
      if (c.city) map.set(normalizeCity(c.city), c.city);
    });
    return [...map.entries()].sort((a, b) => a[1].localeCompare(b[1], "tr"));
  }, [customers]);

  const filtered = useMemo(() => {
    // eslesir(): Türkçe karakter ve büyük/küçük harf duyarsız.
    // Düz toLowerCase() yetmiyordu — "YILMAZ".toLowerCase() "yilmaz" verir,
    // kullanıcı "yılmaz" yazınca kayıt bulunamıyordu.
    const q = query.trim();
    return customers.filter((c) => {
      if (cityFilter && normalizeCity(c.city) !== cityFilter) return false;
      if (!q) return true;
      return eslesir(
        q,
        customerTitle(c), c.company, c.firstName, c.lastName, c.email,
        c.phone, c.addr1, c.addr2, c.city, c.district
      );
    });
  }, [customers, query, cityFilter]);

  const selected = customers.find((c) => c.id === selectedId) || null;

  // Şehir bazlı sayaçlar (Ankara/İstanbul ayrımı için)
  const cityCounts = useMemo(() => {
    const m = new Map<string, number>();
    customers.forEach((c) => {
      const k = normalizeCity(c.city) || "-";
      m.set(k, (m.get(k) || 0) + 1);
    });
    return m;
  }, [customers]);

  function startNew() {
    setForm({ ...BOS });
    setEditing(true);
    setMsg("");
    setErr("");
  }

  function startEdit(c: Customer) {
    const { id, createdAt, updatedAt, ...rest } = c;
    setForm({ ...rest, id, iskontoPct: rest.iskontoPct ?? "" });
    setEditing(true);
    setMsg("");
    setErr("");
  }

  async function save() {
    if (!form.company.trim() && !form.firstName.trim()) {
      setErr("Firma veya kişi adı gerekli.");
      return;
    }
    setSaving(true);
    setErr("");
    try {
      const res = await fetch("/api/musteriler", {
        method: form.id ? "PUT" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Kaydedilemedi");
      setMsg(form.id ? "Müşteri güncellendi." : "Müşteri eklendi.");
      setEditing(false);
      await load(d.customer?.id);
    } catch (e: any) {
      setErr(e.message || "Bir hata oluştu");
    } finally {
      setSaving(false);
    }
  }

  async function remove(c: Customer) {
    if (!confirm(`${customerTitle(c)} kaydı silinsin mi?`)) return;
    const res = await fetch(`/api/musteriler?id=${encodeURIComponent(c.id)}`, {
      method: "DELETE",
    });
    if (res.ok) {
      setMsg("Müşteri silindi.");
      await load();
    }
  }

  const set = (k: keyof Form) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) =>
    setForm((f) => ({ ...f, [k]: e.target.value }));

  return (
    <div className="lbl-layout" ref={layoutRef}>
      {/* ---- Sol: müşteri listesi ---- */}
      <div style={{ minWidth: 0 }}>
        <div className="card pad-sm">
          <div className="card-head">
            <span className="card-head-icon"><Icon name="users" size={16} /></span>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <h2>Kayıtlı Müşteriler</h2>
              <span className="lbl-count">{filtered.length}</span>
            </div>
            <span className="spacer" />
            <div className="card-head-actions">
              <button className="btn small" onClick={startNew}>+ Yeni Müşteri</button>
            </div>
          </div>

          <input
            placeholder="Ara: firma / kişi / telefon / adres"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />

          {/* Şehir filtresi */}
          <div className="lbl-cities">
            <button
              className={`lbl-city ${cityFilter === "" ? "sel" : ""}`}
              onClick={() => setCityFilter("")}
            >
              Tümü <b>{customers.length}</b>
            </button>
            {cities.map(([key, label]) => (
              <button
                key={key}
                className={`lbl-city ${cityFilter === key ? "sel" : ""}`}
                onClick={() => setCityFilter(key)}
              >
                {label} <b>{cityCounts.get(key) || 0}</b>
              </button>
            ))}
          </div>

          {!blobOk && (
            <div className="notice info" style={{ marginTop: 10 }}>
              Kalıcı depolama yapılandırılmadığı için müşteriler kaydedilemiyor.
            </div>
          )}

          <div className="lbl-list">
            {loading ? (
              <div className="empty" style={{ padding: "22px 12px" }}>Yükleniyor...</div>
            ) : filtered.length === 0 ? (
              <div className="empty" style={{ padding: "22px 12px" }}>
                {customers.length === 0
                  ? "Henüz müşteri eklenmedi. “+ Yeni Müşteri” ile başlayın."
                  : "Bu filtreye uyan müşteri yok."}
              </div>
            ) : (
              filtered.map((c) => (
                <div
                  key={c.id}
                  className={`lbl-item ${selectedId === c.id ? "sel" : ""}`}
                  onClick={() => selectCustomer(c.id)}
                >
                  <div className="lbl-item-info">
                    <div className="lbl-item-title">{customerTitle(c)}</div>
                    <div className="lbl-item-meta">
                      {[c.phone, [c.district, c.city].filter(Boolean).join(" / ")]
                        .filter(Boolean)
                        .join(" · ") || "—"}
                    </div>
                  </div>
                  <span className="lbl-item-actions">
                    <span className={`lbl-branch ${c.branch}`}>
                      {c.branch === "istanbul" ? "İST" : "ANK"}
                    </span>
                    <a
                      className="btn small secondary icon"
                      href={`/musteri?id=${encodeURIComponent(c.id)}`}
                      title="Cari kart: sipariş geçmişi ve bakiye"
                      aria-label="Cari kart: sipariş geçmişi ve bakiye"
                      onClick={(e) => e.stopPropagation()}
                    >
                      <Icon name="credit-card" size={15} />
                    </a>
                    <button
                      className="btn small secondary icon"
                      title="Düzenle"
                      aria-label="Düzenle"
                      onClick={(e) => { e.stopPropagation(); startEdit(c); }}
                    >
                      <Icon name="edit" size={15} />
                    </button>
                    <button
                      className="btn small danger icon"
                      title="Sil"
                      aria-label="Sil"
                      onClick={(e) => { e.stopPropagation(); remove(c); }}
                    >
                      <Icon name="x" size={15} />
                    </button>
                  </span>
                </div>
              ))
            )}
          </div>
        </div>

        {/* ---- Müşteri formu ---- */}
        {editing && (
          <div className="card" style={{ marginTop: 16 }}>
            <div className="card-head">
              <span className="card-head-icon">
                <Icon name={form.id ? "edit" : "plus"} size={16} />
              </span>
              <div>
                <h2>{form.id ? "Müşteriyi Düzenle" : "Yeni Müşteri"}</h2>
              </div>
            </div>
            <div className="rw-grid2">
              <div style={{ gridColumn: "1 / -1" }}>
                <label>Firma / Ünvan</label>
                <input value={form.company} onChange={set("company")} placeholder="Örn. Yılmaz Çerçeve Ltd." />
              </div>
              <div>
                <label>Ad</label>
                <input value={form.firstName} onChange={set("firstName")} />
              </div>
              <div>
                <label>Soyad</label>
                <input value={form.lastName} onChange={set("lastName")} />
              </div>
              <div>
                <label>Telefon</label>
                <input value={form.phone} onChange={set("phone")} placeholder="05xx xxx xx xx" />
              </div>
              <div>
                <label>E-posta</label>
                <input value={form.email} onChange={set("email")} />
              </div>
              <div style={{ gridColumn: "1 / -1" }}>
                <label>Adres Satırı 1</label>
                <input value={form.addr1} onChange={set("addr1")} placeholder="Mah. Cad. No:X D:Y" />
              </div>
              <div style={{ gridColumn: "1 / -1" }}>
                <label>Adres Satırı 2 (opsiyonel)</label>
                <input value={form.addr2} onChange={set("addr2")} />
              </div>
              <div>
                <label>İl / Şehir</label>
                <input value={form.city} onChange={set("city")} placeholder="Şehir girin" />
              </div>
              <div>
                <label>İlçe</label>
                <input value={form.district} onChange={set("district")} />
              </div>
              <div>
                <label>Posta Kodu</label>
                <input value={form.postalCode} onChange={set("postalCode")} />
              </div>
              <div>
                <label>Ülke</label>
                <input value={form.country} onChange={set("country")} />
              </div>
              <div>
                <label>Gönderici Şube</label>
                <select value={form.branch} onChange={set("branch")}>
                  <option value="ankara">{BRANCHES.ankara.label}</option>
                  <option value="istanbul">{BRANCHES.istanbul.label}</option>
                </select>
              </div>
              <div>
                <label>Bayi İskontosu (%)</label>
                <input
                  type="text"
                  inputMode="decimal"
                  value={String(form.iskontoPct ?? "")}
                  onChange={set("iskontoPct")}
                  placeholder="0"
                  title="Bu müşteriye özel iskonto — sipariş formunda müşteri seçilince genel iskonto alanına otomatik yazılır"
                />
              </div>
              <div>
                <label>Not (etikette küçük punto)</label>
                <input value={form.note} onChange={set("note")} />
              </div>
            </div>
            <div className="row" style={{ marginTop: 16 }}>
              <button className="btn" disabled={saving} onClick={save}>
                {saving ? "Kaydediliyor..." : form.id ? "Güncelle" : "Kaydet"}
              </button>
              <button className="btn secondary" onClick={() => setEditing(false)}>
                Vazgeç
              </button>
            </div>
            {err && <div className="notice err">{err}</div>}
          </div>
        )}
        {msg && !editing && <div className="notice ok">{msg}</div>}
      </div>

      {/* ---- Sağ: etiket önizleme ---- */}
      <aside
        ref={previewRef}
        style={{ minWidth: 0, scrollMarginTop: "calc(var(--topbar-h) + 12px)" }}
      >
        <div
          className="card pad-sm"
          style={{ position: "sticky", top: "calc(var(--topbar-h) + 16px)" }}
        >
          <div className="card-head">
            <span className="card-head-icon"><Icon name="tag" size={16} /></span>
            <div>
              <h2>Kargo Etiketi</h2>
              <span className="card-head-sub">150 × 100 mm</span>
            </div>
          </div>

          {selected ? (
            <>
              <LabelPreview c={selected} />

              {/* Gönderici şubesi hızlı değiştirme */}
              <div style={{ marginTop: 12 }}>
                <label>Gönderici Şube</label>
                <div className="seg" style={{ width: "100%", display: "flex" }}>
                  {(["ankara", "istanbul"] as Branch[]).map((b) => (
                    <button
                      key={b}
                      type="button"
                      className={selected.branch === b ? "active" : ""}
                      style={{ flex: 1 }}
                      onClick={async () => {
                        const upd = { ...selected, branch: b };
                        setCustomers((list) =>
                          list.map((x) => (x.id === upd.id ? upd : x))
                        );
                        await fetch("/api/musteriler", {
                          method: "PUT",
                          headers: { "Content-Type": "application/json" },
                          body: JSON.stringify(upd),
                        });
                      }}
                    >
                      {b === "ankara" ? "Ankara" : "İstanbul"}
                    </button>
                  ))}
                </div>
                <p className="muted" style={{ fontSize: 11.5, marginTop: 6 }}>
                  Etiketin üstündeki gönderici adresi seçilen şubeye göre değişir.
                </p>
              </div>

              <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "flex-end", minWidth: 0 }}>
                <div style={{ flex: "0 0 90px", maxWidth: "100%" }}>
                  <label>Adet</label>
                  <input
                    type="number"
                    min="1"
                    max="50"
                    value={copies}
                    onChange={(e) => setCopies(e.target.value)}
                  />
                </div>
                <a
                  className="btn"
                  style={{ flex: 1, minWidth: 0 }}
                  href={`/api/etiket/pdf?id=${encodeURIComponent(selected.id)}&adet=${Math.max(1, parseInt(copies) || 1)}`}
                >
                  <Icon name="download" size={16} /> Etiket PDF
                </a>
              </div>
            </>
          ) : (
            <div className="empty" style={{ padding: "22px 12px" }}>
              <div className="empty-icon"><Icon name="tag" size={22} /></div>
              Etiketi görmek için soldaki listeden bir müşteri seçin.
            </div>
          )}
        </div>
      </aside>
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
      {/* Uzun firma adı / adres etiket kutusundan taşmasın — kâğıtta da kesilir */}
      <div className="lbl-to" style={{ minHeight: 0, overflow: "hidden", overflowWrap: "anywhere" }}>
        <div className="lbl-to-title">Alıcı</div>
        <div className="lbl-to-name">{customerTitle(c)}</div>
        {(c.phone || c.email) && (
          <div className="lbl-to-contact">
            {[c.phone && `Tel: ${c.phone}`, c.email && `E-posta: ${c.email}`]
              .filter(Boolean)
              .join("  •  ")}
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
