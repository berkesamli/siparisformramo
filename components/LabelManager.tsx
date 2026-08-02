"use client";

// Müşteri defteri + 150×100 mm kargo etiketi.
// "etiket" uygulamasının portu: kayıtlı müşteriler, arama, şehir filtresi,
// müşteri ekle/düzenle/sil, gönderici şubesine göre değişen etiket önizlemesi.

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  BRANCHES,
  branchInfo,
  customerTitle,
  normalizeCity,
  type Branch,
  type Customer,
} from "@/lib/customers";

const BOS: Omit<Customer, "id" | "createdAt" | "updatedAt"> = {
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
    const q = query.trim().toLowerCase();
    return customers.filter((c) => {
      if (cityFilter && normalizeCity(c.city) !== cityFilter) return false;
      if (!q) return true;
      const hay = [
        customerTitle(c), c.company, c.firstName, c.lastName, c.email,
        c.phone, c.addr1, c.addr2, c.city, c.district,
      ].join(" ").toLowerCase();
      return hay.includes(q);
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
    setForm({ ...rest, id });
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
    <div className="lbl-layout">
      {/* ---- Sol: müşteri listesi ---- */}
      <div>
        <div className="card" style={{ padding: 16 }}>
          <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
            <h2 style={{ margin: 0, fontSize: 16 }}>Kayıtlı Müşteriler</h2>
            <span className="lbl-count">{filtered.length}</span>
            <span style={{ flex: 1 }} />
            <button className="btn small" onClick={startNew}>+ Yeni Müşteri</button>
          </div>

          <input
            style={{ marginTop: 12 }}
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
              <p style={{ color: "var(--muted)", padding: 10 }}>Yükleniyor...</p>
            ) : filtered.length === 0 ? (
              <p style={{ color: "var(--muted)", padding: 10 }}>
                {customers.length === 0
                  ? "Henüz müşteri eklenmedi. “+ Yeni Müşteri” ile başlayın."
                  : "Bu filtreye uyan müşteri yok."}
              </p>
            ) : (
              filtered.map((c) => (
                <div
                  key={c.id}
                  className={`lbl-item ${selectedId === c.id ? "sel" : ""}`}
                  onClick={() => setSelectedId(c.id)}
                >
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div className="lbl-item-title">{customerTitle(c)}</div>
                    <div className="lbl-item-meta">
                      {[c.phone, [c.district, c.city].filter(Boolean).join(" / ")]
                        .filter(Boolean)
                        .join(" · ") || "—"}
                    </div>
                  </div>
                  <span className={`lbl-branch ${c.branch}`}>
                    {c.branch === "istanbul" ? "İST" : "ANK"}
                  </span>
                  <button
                    className="btn small secondary"
                    onClick={(e) => { e.stopPropagation(); startEdit(c); }}
                  >
                    ✏️
                  </button>
                  <button
                    className="btn small danger"
                    onClick={(e) => { e.stopPropagation(); remove(c); }}
                  >
                    🗑
                  </button>
                </div>
              ))
            )}
          </div>
        </div>

        {/* ---- Müşteri formu ---- */}
        {editing && (
          <div className="card" style={{ marginTop: 14 }}>
            <h2 style={{ marginTop: 0 }}>
              {form.id ? "🖊️ Müşteriyi Düzenle" : "➕ Yeni Müşteri"}
            </h2>
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
                <input value={form.city} onChange={set("city")} placeholder="Ankara" />
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
                <label>Not (etikette küçük punto)</label>
                <input value={form.note} onChange={set("note")} />
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, marginTop: 16, flexWrap: "wrap" }}>
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
      <aside>
        <div className="card" style={{ position: "sticky", top: 90, padding: 18 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
            <b style={{ fontSize: 15 }}>Kargo Etiketi</b>
            <span style={{ fontSize: 11.5, color: "var(--muted)" }}>150 × 100 mm</span>
          </div>

          {selected ? (
            <>
              <LabelPreview c={selected} />

              {/* Gönderici şubesi hızlı değiştirme */}
              <div style={{ marginTop: 12 }}>
                <label>Gönderici Şube</label>
                <div className="rw-toggle" style={{ width: "100%" }}>
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
                <p style={{ fontSize: 11.5, color: "var(--muted)", marginTop: 6 }}>
                  Etiketin üstündeki gönderici adresi seçilen şubeye göre değişir.
                </p>
              </div>

              <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "flex-end" }}>
                <div style={{ width: 90 }}>
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
                  style={{ flex: 1, justifyContent: "center" }}
                  href={`/api/etiket/pdf?id=${encodeURIComponent(selected.id)}&adet=${Math.max(1, parseInt(copies) || 1)}`}
                >
                  ⬇ Etiket PDF
                </a>
              </div>
            </>
          ) : (
            <p style={{ color: "var(--muted)", fontSize: 13.5 }}>
              Etiketi görmek için soldaki listeden bir müşteri seçin.
            </p>
          )}
        </div>
      </aside>
    </div>
  );
}

/* ---- Ekrandaki etiket önizlemesi (PDF ile aynı düzen) ---- */
function LabelPreview({ c }: { c: Customer }) {
  const b = branchInfo(c.branch);
  const lines = [
    c.addr1,
    c.addr2,
    [c.district, c.city].filter(Boolean).join(" / "),
    [c.postalCode, c.country].filter(Boolean).join(" "),
  ].filter(Boolean);

  return (
    <div className="lbl-preview">
      <div className="lbl-hdr">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/logo.png" alt="Olga Çerçeve" />
        <div>
          <div className="lbl-firm">{b.name}</div>
          <div className="lbl-tel">{b.cityTel}</div>
          <div className="lbl-addr">{b.addr1}, {b.addr2}</div>
          <div className="lbl-addr">{b.website}</div>
        </div>
      </div>
      <div className="lbl-to">
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
