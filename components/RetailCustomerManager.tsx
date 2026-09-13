"use client";

// Perakende müşteri defteri yönetimi — listele/ara, ekle, düzenle, sil.
// Kayıtlar retail-customers/ altında durur (etiket defterinden ayrı).

import { useCallback, useEffect, useMemo, useState } from "react";
import Icon from "@/components/shell/Icon";
import type { RetailCustomer } from "@/lib/retail-customers";
import { eslesir } from "@/lib/search-norm";

const BOS = { id: "", name: "", phone: "", email: "", address: "", note: "" };

export default function RetailCustomerManager({
  initialQuery,
}: {
  /** Arama kutusunu tohumlayan başlangıç sorgusu (sayfa ?q= ile geçirir). */
  initialQuery?: string;
} = {}) {
  const [customers, setCustomers] = useState<RetailCustomer[]>([]);
  const [loading, setLoading] = useState(true);
  const [blobOk, setBlobOk] = useState(true);
  const [query, setQuery] = useState(initialQuery ?? "");
  const [form, setForm] = useState({ ...BOS });
  const [editing, setEditing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState("");
  const [err, setErr] = useState("");

  // Sayfa yeni bir ?q= ile gelirse arama kutusu da onu izlesin.
  useEffect(() => {
    if (initialQuery !== undefined) setQuery(initialQuery);
  }, [initialQuery]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/perakende/musteriler");
      const d = await res.json();
      if (res.ok) {
        setCustomers(d.customers || []);
        setBlobOk(d.blob !== false);
      }
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const filtered = useMemo(() => {
    const q = query.trim();
    if (!q) return customers;
    return customers.filter((c) => eslesir(q, c.name, c.phone, c.email, c.address));
  }, [customers, query]);

  const set =
    (k: keyof typeof BOS) => (e: React.ChangeEvent<HTMLInputElement>) =>
      setForm((f) => ({ ...f, [k]: e.target.value }));

  async function save() {
    if (!form.name.trim()) {
      setErr("Müşteri adı gerekli.");
      return;
    }
    setSaving(true);
    setErr("");
    try {
      const res = await fetch("/api/perakende/musteriler", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Kaydedilemedi");
      setMsg(form.id ? "Müşteri güncellendi." : "Müşteri eklendi.");
      setEditing(false);
      setForm({ ...BOS });
      await load();
    } catch (e: any) {
      setErr(e.message || "Bir hata oluştu");
    } finally {
      setSaving(false);
    }
  }

  async function remove(c: RetailCustomer) {
    if (!confirm(`${c.name} kaydı silinsin mi? (Geçmiş siparişleri silinmez.)`)) return;
    const res = await fetch(`/api/perakende/musteriler?id=${encodeURIComponent(c.id)}`, {
      method: "DELETE",
    });
    if (res.ok) {
      setMsg("Müşteri silindi.");
      await load();
    }
  }

  // Uzun serbest metin (adres / not) satırı şişirmesin: tek satır + üç nokta,
  // tam metin title'da. Tablo telefonda .table-wrap içinde yatay kayar.
  const ellipsis: React.CSSProperties = {
    fontSize: 13,
    maxWidth: 240,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  };

  return (
    <div className="card">
      <div className="row no-print" style={{ marginBottom: 12 }}>
        <div style={{ flex: 1, minWidth: "min(220px, 100%)", position: "relative" }}>
          <span
            aria-hidden
            style={{
              position: "absolute",
              left: 12,
              top: "50%",
              transform: "translateY(-50%)",
              color: "var(--muted)",
              display: "inline-flex",
              pointerEvents: "none",
            }}
          >
            <Icon name="search" size={15} />
          </span>
          <input
            style={{ paddingLeft: 36 }}
            placeholder="Ara: ad / telefon / e-posta / adres"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            aria-label="Müşteri ara"
          />
        </div>
        {/* Rozet + birincil buton birlikte kalır; telefonda ikinci satıra
            düştüklerinde sağa yaslanır, arama üstte tam genişlik alır. */}
        <div className="row" style={{ marginLeft: "auto", gap: 8, flexWrap: "nowrap" }}>
          <span className="badge">{filtered.length} kayıt</span>
          <button
            className="btn small"
            onClick={() => {
              setForm({ ...BOS });
              setEditing(true);
              setMsg("");
              setErr("");
            }}
          >
            + Yeni Müşteri
          </button>
        </div>
      </div>

      {!blobOk && (
        <div className="notice info" style={{ marginBottom: 12 }}>
          Kalıcı depolama yapılandırılmadığı için kayıtlar saklanamıyor.
        </div>
      )}
      {msg && <div className="notice ok" style={{ marginBottom: 12 }}>{msg}</div>}
      {err && <div className="notice err" style={{ marginBottom: 12 }}>{err}</div>}

      {editing && (
        <div
          className="no-print"
          style={{
            marginBottom: 16,
            padding: 16,
            background: "var(--surface-2)",
            border: "1px solid var(--border)",
            borderRadius: "var(--radius-sm)",
          }}
        >
          <div className="row" style={{ marginBottom: 12, gap: 8 }}>
            <span className="card-head-icon" style={{ width: 30, height: 30 }}>
              <Icon name={form.id ? "edit" : "plus"} size={15} />
            </span>
            <h3 style={{ margin: 0 }}>{form.id ? "Müşteriyi Düzenle" : "Yeni Müşteri"}</h3>
          </div>
          <div className="grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(min(200px, 100%), 1fr))" }}>
            <div>
              <label>Ad Soyad *</label>
              <input value={form.name} onChange={set("name")} />
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
              <label>Adres</label>
              <input value={form.address} onChange={set("address")} />
            </div>
            <div style={{ gridColumn: "1 / -1" }}>
              <label>Not</label>
              <input value={form.note} onChange={set("note")} placeholder="örn. köşedeki galeri, pazartesi kapalı" />
            </div>
          </div>
          <div className="row" style={{ marginTop: 12 }}>
            <button className="btn small" disabled={saving} onClick={save}>
              {saving ? "Kaydediliyor…" : form.id ? "Güncelle" : "Kaydet"}
            </button>
            <button className="btn small secondary" onClick={() => setEditing(false)}>
              Vazgeç
            </button>
          </div>
        </div>
      )}

      {loading ? (
        <div className="empty" style={{ padding: "22px 12px" }}>Yükleniyor…</div>
      ) : filtered.length === 0 ? (
        <div className="empty" style={{ padding: "22px 12px" }}>
          <div className="empty-icon"><Icon name="users" size={22} /></div>
          {customers.length === 0
            ? "Henüz perakende müşterisi yok — ilk sipariş kaydedilince müşteri buraya kendiliğinden eklenir."
            : "Bu aramaya uyan müşteri yok."}
        </div>
      ) : (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Ad Soyad</th>
                <th>Telefon</th>
                <th>E-posta</th>
                <th>Adres</th>
                <th>Not</th>
                <th className="no-print">İşlem</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((c) => (
                <tr key={c.id}>
                  <td style={{ fontWeight: 600, whiteSpace: "nowrap" }}>{c.name}</td>
                  <td style={{ whiteSpace: "nowrap" }}>{c.phone || "—"}</td>
                  <td style={{ whiteSpace: "nowrap" }}>{c.email || "—"}</td>
                  <td style={ellipsis} title={c.address || undefined}>{c.address || "—"}</td>
                  <td style={ellipsis} title={c.note || undefined}>{c.note || "—"}</td>
                  <td className="no-print" style={{ whiteSpace: "nowrap" }}>
                    <span style={{ display: "inline-flex", gap: 6 }}>
                      <button
                        className="btn small secondary icon"
                        title="Düzenle"
                        aria-label="Düzenle"
                        onClick={() => {
                          setForm({
                            id: c.id,
                            name: c.name,
                            phone: c.phone,
                            email: c.email,
                            address: c.address,
                            note: c.note,
                          });
                          setEditing(true);
                          setMsg("");
                          setErr("");
                          window.scrollTo({ top: 0, behavior: "smooth" });
                        }}
                      >
                        <Icon name="edit" size={15} />
                      </button>
                      <button
                        className="btn small danger icon"
                        title="Sil"
                        aria-label="Sil"
                        onClick={() => remove(c)}
                      >
                        <Icon name="x" size={15} />
                      </button>
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
