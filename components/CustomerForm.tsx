"use client";

// Müşteri ekle / düzenle modalı — Müşteriler sayfası, müşteri kartı ve
// kargo etiketi sayfası aynı formu kullanır. Telefonda tam ekran açılır.
// Mikro eşleştirme alanları buradan gönderilmez; sunucu mevcut değeri korur.

import { useEffect, useState } from "react";
import { createPortal } from "react-dom";
import Icon from "@/components/shell/Icon";
import { BRANCHES, type Customer } from "@/lib/customers";

interface FormState {
  company: string; firstName: string; lastName: string; phone: string; email: string;
  addr1: string; addr2: string; city: string; district: string; postalCode: string; country: string;
  branch: "ankara" | "istanbul"; iskontoPct: string; note: string;
}

const BOS: FormState = {
  company: "", firstName: "", lastName: "", phone: "", email: "",
  addr1: "", addr2: "", city: "", district: "", postalCode: "", country: "Türkiye",
  branch: "ankara", iskontoPct: "", note: "",
};

function fromCustomer(c: Customer): FormState {
  return {
    company: c.company || "", firstName: c.firstName || "", lastName: c.lastName || "",
    phone: c.phone || "", email: c.email || "",
    addr1: c.addr1 || "", addr2: c.addr2 || "", city: c.city || "", district: c.district || "",
    postalCode: c.postalCode || "", country: c.country || "Türkiye",
    branch: c.branch === "istanbul" ? "istanbul" : "ankara",
    iskontoPct: c.iskontoPct ? String(c.iskontoPct) : "",
    note: c.note || "",
  };
}

export default function CustomerForm({
  initial,
  onSaved,
  onClose,
}: {
  initial?: Customer | null;
  onSaved: (c: Customer) => void;
  onClose: () => void;
}) {
  const [form, setForm] = useState<FormState>(() => (initial ? fromCustomer(initial) : { ...BOS }));
  const [saving, setSaving] = useState(false);
  const [err, setErr] = useState("");

  // Esc ile kapat; arkadaki sayfa kaymasın
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    document.addEventListener("keydown", onKey);
    const prev = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.removeEventListener("keydown", onKey); document.body.style.overflow = prev; };
  }, [onClose]);

  const set = (k: keyof FormState) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) =>
    setForm((f) => ({ ...f, [k]: e.target.value }));

  async function save() {
    if (!form.company.trim() && !form.firstName.trim()) {
      setErr("Firma adı ya da kişi adı gerekli.");
      return;
    }
    setSaving(true); setErr("");
    try {
      const res = await fetch("/api/musteriler", {
        method: initial ? "PUT" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(initial ? { ...form, id: initial.id } : form),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Kaydedilemedi.");
      onSaved(d.customer as Customer);
    } catch (e: any) {
      setErr(e?.message || "Bir hata oluştu.");
    } finally {
      setSaving(false);
    }
  }

  if (typeof document === "undefined") return null;
  return createPortal(
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-card cf-card" role="dialog" aria-modal="true" aria-label={initial ? "Müşteriyi düzenle" : "Yeni müşteri"} onClick={(e) => e.stopPropagation()}>
        <div className="cf-head">
          <span className="card-head-icon"><Icon name={initial ? "edit" : "plus"} size={16} /></span>
          <h2>{initial ? "Müşteriyi Düzenle" : "Yeni Müşteri"}</h2>
          <button type="button" className="btn ghost icon small" onClick={onClose} aria-label="Kapat"><Icon name="x" size={18} /></button>
        </div>
        <form
          className="cf-body"
          onSubmit={(e) => { e.preventDefault(); save(); }}
        >
          <div className="cf-grid">
            <div className="cf-sec">Kimlik</div>
            <div className="span2">
              <label>Firma / Ünvan</label>
              <input value={form.company} onChange={set("company")} placeholder="Örn. Yılmaz Çerçeve Ltd." autoFocus />
            </div>
            <div>
              <label>Yetkili Adı</label>
              <input value={form.firstName} onChange={set("firstName")} placeholder="Ad" />
            </div>
            <div>
              <label>Soyadı</label>
              <input value={form.lastName} onChange={set("lastName")} placeholder="Soyad" />
            </div>

            <div className="cf-sec">İletişim</div>
            <div>
              <label>Telefon</label>
              <input type="tel" inputMode="tel" value={form.phone} onChange={set("phone")} placeholder="05xx xxx xx xx" />
            </div>
            <div>
              <label>E-posta</label>
              <input type="email" inputMode="email" value={form.email} onChange={set("email")} placeholder="ornek@firma.com" />
            </div>

            <div className="cf-sec">Adres (kargo etiketi)</div>
            <div className="span2">
              <label>Adres Satırı 1</label>
              <input value={form.addr1} onChange={set("addr1")} placeholder="Mah. Cad. No:X D:Y" />
            </div>
            <div className="span2">
              <label>Adres Satırı 2 (opsiyonel)</label>
              <input value={form.addr2} onChange={set("addr2")} />
            </div>
            <div>
              <label>İl</label>
              <input value={form.city} onChange={set("city")} placeholder="Örn. Ankara" />
            </div>
            <div>
              <label>İlçe</label>
              <input value={form.district} onChange={set("district")} />
            </div>
            <div>
              <label>Posta Kodu</label>
              <input inputMode="numeric" value={form.postalCode} onChange={set("postalCode")} />
            </div>
            <div>
              <label>Ülke</label>
              <input value={form.country} onChange={set("country")} />
            </div>

            <div className="cf-sec">Satış</div>
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
                type="text" inputMode="decimal"
                value={form.iskontoPct} onChange={set("iskontoPct")} placeholder="0"
                title="Sipariş formunda bu müşteri seçilince genel iskonto alanına otomatik yazılır"
              />
            </div>
            <div className="span2">
              <label>Not (etikette küçük punto)</label>
              <input value={form.note} onChange={set("note")} placeholder="Örn. Kapıda ara, 2. kat" />
            </div>
          </div>
        </form>
        <div className="cf-foot">
          {err && <div className="notice err">{err}</div>}
          <button type="button" className="btn secondary" onClick={onClose} disabled={saving}>Vazgeç</button>
          <button type="button" className="btn" onClick={save} disabled={saving}>
            <Icon name="check-circle" size={16} /> {saving ? "Kaydediliyor…" : initial ? "Güncelle" : "Kaydet"}
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}
