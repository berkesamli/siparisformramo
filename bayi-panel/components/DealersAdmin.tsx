"use client";

// Olga yönetici paneli — bayi tanımlama, abonelik durumu, şifre sıfırlama.

import { useCallback, useEffect, useState } from "react";
import { SUBSCRIPTION_LABELS, type SubscriptionStatus } from "@/data/subscription";
import type { PublicDealer } from "@/lib/dealers";

type DealerRow = PublicDealer & {
  stats?: { orders: number; orders30: number; revenue30: number; lastOrderAt: string | null };
};

const fmt0 = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");

const STATUS_COLORS: Record<SubscriptionStatus, string> = {
  aktif: "#067a55",
  muaf: "#1d4ed8",
  odeme_bekliyor: "#b45309",
  askida: "#b91c1c",
};

const EMPTY_FORM = {
  name: "", username: "", password: "", contactName: "", phone: "", email: "", city: "", website: "", address: "",
  status: "aktif" as SubscriptionStatus, paidUntil: "", monthlyFee: "", note: "",
};

export default function DealersAdmin() {
  const [dealers, setDealers] = useState<DealerRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [blob, setBlob] = useState(true);
  const [defaults, setDefaults] = useState({ fee: 0, threshold: 0 });
  const [showNew, setShowNew] = useState(false);
  const [form, setForm] = useState({ ...EMPTY_FORM });
  const [open, setOpen] = useState<string | null>(null);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const [busy, setBusy] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/yonetim/bayiler");
      const d = await res.json();
      if (res.ok) {
        setDealers(d.dealers || []);
        setBlob(d.blob !== false);
        setDefaults(d.defaults || { fee: 0, threshold: 0 });
      }
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  async function api(method: string, body?: unknown, qs = "") {
    setBusy(true);
    setMsg(null);
    try {
      const res = await fetch(`/api/yonetim/bayiler${qs}`, {
        method,
        headers: { "Content-Type": "application/json" },
        body: body ? JSON.stringify(body) : undefined,
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "İşlem başarısız");
      await load();
      return true;
    } catch (e: any) {
      setMsg({ ok: false, text: e.message });
      return false;
    } finally {
      setBusy(false);
    }
  }

  async function createDealer() {
    const ok = await api("POST", {
      ...form,
      subscription: { status: form.status, paidUntil: form.paidUntil, monthlyFee: form.monthlyFee, note: form.note },
    });
    if (ok) {
      setMsg({ ok: true, text: `${form.name} bayisi oluşturuldu. Giriş: ${form.username}` });
      setForm({ ...EMPTY_FORM });
      setShowNew(false);
    }
  }

  async function patchDealer(slug: string, body: Record<string, unknown>) {
    const ok = await api("PATCH", { slug, ...body });
    if (ok) setMsg({ ok: true, text: "Güncellendi." });
  }

  async function resetPassword(d: DealerRow) {
    const np = prompt(`${d.name} için yeni şifre (en az 6 karakter):`);
    if (!np) return;
    const ok = await api("PATCH", { slug: d.slug, newPassword: np });
    if (ok) setMsg({ ok: true, text: `${d.name} şifresi güncellendi.` });
  }

  async function removeDealer(d: DealerRow) {
    if (!confirm(`${d.name} bayisi silinsin mi? Hesap kapanır; sipariş kayıtları arşivde kalır.`)) return;
    const ok = await api("DELETE", undefined, `?slug=${encodeURIComponent(d.slug)}`);
    if (ok) setMsg({ ok: true, text: `${d.name} silindi.` });
  }

  const f = (k: keyof typeof EMPTY_FORM) => ({
    value: form[k] as string,
    onChange: (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => setForm({ ...form, [k]: e.target.value }),
  });

  return (
    <div>
      {!blob && (
        <div className="notice err">Kalıcı depolama (Vercel Blob) yapılandırılmamış; bayi kaydı yapılamaz.</div>
      )}
      {msg && <div className={`notice ${msg.ok ? "ok" : "err"}`}>{msg.text}</div>}

      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 14 }}>
        <button className="btn" onClick={() => setShowNew(!showNew)}>{showNew ? "Kapat" : "➕ Yeni Bayi"}</button>
        <span style={{ flex: 1 }} />
        <span style={{ fontSize: 13, color: "var(--muted)" }}>
          {defaults.fee > 0 && <>Aylık ücret: ₺{fmt0(defaults.fee)} · </>}
          {defaults.threshold > 0 && <>Muafiyet eşiği: aylık ₺{fmt0(defaults.threshold)} alım</>}
        </span>
      </div>

      {showNew && (
        <div className="card">
          <h2 style={{ marginTop: 0 }}>Yeni Bayi</h2>
          <div className="rw-grid2">
            <div><label>Firma Adı *</label><input {...f("name")} placeholder="Örn. Ankara Çerçeve Evi" /></div>
            <div><label>Yetkili</label><input {...f("contactName")} /></div>
            <div><label>Kullanıcı Adı *</label><input {...f("username")} placeholder="küçük harf, boşluksuz" /></div>
            <div><label>Şifre * (en az 6)</label><input {...f("password")} type="text" /></div>
            <div><label>Telefon</label><input {...f("phone")} /></div>
            <div><label>E-posta</label><input {...f("email")} type="email" /></div>
            <div><label>Şehir</label><input {...f("city")} /></div>
            <div><label>Web sitesi</label><input {...f("website")} /></div>
            <div style={{ gridColumn: "1 / -1" }}><label>Adres</label><input {...f("address")} /></div>
            <div>
              <label>Abonelik Durumu</label>
              <select {...f("status")}>
                {Object.entries(SUBSCRIPTION_LABELS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
            </div>
            <div><label>Ödenmiş Son Tarih</label><input {...f("paidUntil")} type="date" /></div>
            <div><label>Aylık Ücret (₺)</label><input {...f("monthlyFee")} type="number" placeholder={defaults.fee ? String(defaults.fee) : ""} /></div>
            <div><label>Not</label><input {...f("note")} placeholder="örn. 3 aylık alım 62.000 TL — muaf" /></div>
          </div>
          <button className="btn" style={{ marginTop: 14 }} disabled={busy || !blob} onClick={createDealer}>
            Bayiyi Oluştur
          </button>
        </div>
      )}

      {loading ? (
        <p style={{ color: "var(--muted)" }}>Yükleniyor...</p>
      ) : dealers.length === 0 ? (
        <div className="card" style={{ textAlign: "center", color: "var(--muted)" }}>Henüz bayi tanımlanmadı.</div>
      ) : (
        dealers.map((d) => (
          <div className="card" key={d.slug} style={{ padding: 16, marginBottom: 12, opacity: d.active ? 1 : 0.6 }}>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
              <strong style={{ color: "var(--brand)", fontSize: 15 }}>{d.name}</strong>
              <span style={{ fontSize: 13, color: "var(--text-2)" }}>
                {d.city && `${d.city} · `}{d.phone}{d.contactName && ` · ${d.contactName}`}
              </span>
              <span style={{ flex: 1 }} />
              {d.stats && (
                <span style={{ fontSize: 13, color: "var(--muted)" }}>
                  30 gün: <strong>{d.stats.orders30}</strong> sipariş · ₺{fmt0(d.stats.revenue30)} · toplam {d.stats.orders}
                </span>
              )}
              <select
                style={{ width: 210, fontWeight: 600, color: STATUS_COLORS[d.subscription.status] }}
                value={d.subscription.status}
                disabled={busy}
                onChange={(e) => patchDealer(d.slug, { subscription: { status: e.target.value } })}
              >
                {Object.entries(SUBSCRIPTION_LABELS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
              <button className="btn small secondary" onClick={() => setOpen(open === d.slug ? null : d.slug)}>
                {open === d.slug ? "Kapat" : "Düzenle"}
              </button>
            </div>

            {open === d.slug && (
              <DealerEditor d={d} busy={busy} onSave={(body) => patchDealer(d.slug, body)} onReset={() => resetPassword(d)} onDelete={() => removeDealer(d)} />
            )}
          </div>
        ))
      )}
    </div>
  );
}

function DealerEditor({
  d, busy, onSave, onReset, onDelete,
}: {
  d: DealerRow;
  busy: boolean;
  onSave: (body: Record<string, unknown>) => void;
  onReset: () => void;
  onDelete: () => void;
}) {
  const [v, setV] = useState({
    name: d.name, contactName: d.contactName || "", username: d.username, phone: d.phone || "", email: d.email || "",
    city: d.city || "", website: d.website || "", address: d.address || "", active: d.active,
    paidUntil: d.subscription.paidUntil || "", monthlyFee: d.subscription.monthlyFee ?? "", note: d.subscription.note || "",
  });
  const f = (k: keyof typeof v) => ({
    value: String(v[k] ?? ""),
    onChange: (e: React.ChangeEvent<HTMLInputElement>) => setV({ ...v, [k]: e.target.value }),
  });
  return (
    <div style={{ marginTop: 12, borderTop: "1px solid var(--border)", paddingTop: 12 }}>
      <div className="rw-grid2">
        <div><label>Firma Adı</label><input {...f("name")} /></div>
        <div><label>Yetkili</label><input {...f("contactName")} /></div>
        <div><label>Kullanıcı Adı</label><input {...f("username")} /></div>
        <div><label>Telefon</label><input {...f("phone")} /></div>
        <div><label>E-posta</label><input {...f("email")} /></div>
        <div><label>Şehir</label><input {...f("city")} /></div>
        <div><label>Web sitesi</label><input {...f("website")} /></div>
        <div><label>Adres</label><input {...f("address")} /></div>
        <div><label>Ödenmiş Son Tarih</label><input {...f("paidUntil")} type="date" /></div>
        <div><label>Aylık Ücret (₺)</label><input {...f("monthlyFee")} type="number" /></div>
        <div style={{ gridColumn: "1 / -1" }}><label>Abonelik Notu</label><input {...f("note")} /></div>
        <div>
          <label>Hesap</label>
          <select value={v.active ? "1" : "0"} onChange={(e) => setV({ ...v, active: e.target.value === "1" })}>
            <option value="1">Aktif</option>
            <option value="0">Pasif (giriş yapamaz)</option>
          </select>
        </div>
      </div>
      <div style={{ fontSize: 12.5, color: "var(--muted)", marginTop: 8 }}>
        Bayi kodu: {d.slug} · Kayıt: {new Date(d.createdAt).toLocaleDateString("tr-TR")}
        {d.stats?.lastOrderAt && ` · Son sipariş: ${new Date(d.stats.lastOrderAt).toLocaleDateString("tr-TR")}`}
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
        <button
          className="btn"
          disabled={busy}
          onClick={() =>
            onSave({
              name: v.name, contactName: v.contactName, username: v.username, phone: v.phone, email: v.email,
              city: v.city, website: v.website, address: v.address, active: v.active,
              subscription: { paidUntil: v.paidUntil, monthlyFee: v.monthlyFee, note: v.note },
            })
          }
        >
          💾 Kaydet
        </button>
        <button className="btn secondary" disabled={busy} onClick={onReset}>🔑 Şifre Sıfırla</button>
        <span style={{ flex: 1 }} />
        <button className="btn danger small" disabled={busy} onClick={onDelete}>Bayiyi Sil</button>
      </div>
    </div>
  );
}
