"use client";

// Perakende müşteri defteri yönetimi — listele/ara, ekle, düzenle, sil.
// Kayıtlar retail-customers/ altında durur (etiket defterinden ayrı).

import { useCallback, useEffect, useMemo, useState } from "react";
import type { RetailCustomer } from "@/lib/retail-customers";
import { eslesir } from "@/lib/search-norm";

const BOS = { id: "", name: "", phone: "", email: "", address: "", note: "" };

export default function RetailCustomerManager() {
  const [customers, setCustomers] = useState<RetailCustomer[]>([]);
  const [loading, setLoading] = useState(true);
  const [blobOk, setBlobOk] = useState(true);
  const [query, setQuery] = useState("");
  const [form, setForm] = useState({ ...BOS });
  const [editing, setEditing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState("");
  const [err, setErr] = useState("");

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

  return (
    <div className="card">
      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 12 }}>
        <input
          style={{ flex: 1, minWidth: 220 }}
          placeholder="Ara: ad / telefon / e-posta / adres"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <span style={{ color: "var(--muted)", fontSize: 13 }}>{filtered.length} kayıt</span>
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

      {!blobOk && (
        <div className="notice info" style={{ marginBottom: 12 }}>
          Kalıcı depolama yapılandırılmadığı için kayıtlar saklanamıyor.
        </div>
      )}
      {msg && <div className="notice ok" style={{ marginBottom: 12 }}>{msg}</div>}
      {err && <div className="notice err" style={{ marginBottom: 12 }}>{err}</div>}

      {editing && (
        <div
          className="card"
          style={{ marginBottom: 16, background: "rgba(255,255,255,0.03)" }}
        >
          <div className="grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))" }}>
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
          <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
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
        <p style={{ color: "var(--muted)" }}>Yükleniyor…</p>
      ) : filtered.length === 0 ? (
        <p style={{ color: "var(--muted)" }}>
          {customers.length === 0
            ? "Henüz perakende müşterisi yok — ilk sipariş kaydedilince müşteri buraya kendiliğinden eklenir."
            : "Bu aramaya uyan müşteri yok."}
        </p>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Ad Soyad</th>
                <th>Telefon</th>
                <th>E-posta</th>
                <th>Adres</th>
                <th>Not</th>
                <th>İşlem</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((c) => (
                <tr key={c.id}>
                  <td style={{ fontWeight: 600 }}>{c.name}</td>
                  <td style={{ whiteSpace: "nowrap" }}>{c.phone || "—"}</td>
                  <td>{c.email || "—"}</td>
                  <td style={{ fontSize: 13 }}>{c.address || "—"}</td>
                  <td style={{ fontSize: 13 }}>{c.note || "—"}</td>
                  <td style={{ whiteSpace: "nowrap" }}>
                    <button
                      className="btn small secondary"
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
                      ✏️
                    </button>{" "}
                    <button className="btn small danger" onClick={() => remove(c)}>
                      🗑️
                    </button>
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
