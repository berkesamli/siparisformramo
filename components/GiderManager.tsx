"use client";

// Gider yönetimi — ay/şube filtreli liste + giriş formu.
// Kasa Excel'indeki "ÇIKIŞLAR" bölümünün karşılığı.

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  GIDER_KATEGORILERI,
  GIDER_YONTEM_LABELS,
  type Gider,
  type GiderYontem,
} from "@/lib/gider";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const buAy = () =>
  new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" }).slice(0, 7);

export default function GiderManager() {
  const [ay, setAy] = useState(buAy());
  const [sube, setSube] = useState("");
  const [records, setRecords] = useState<Gider[]>([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [formOpen, setFormOpen] = useState(false);

  // form
  const bugun = new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" });
  const [fDate, setFDate] = useState(bugun);
  const [fBranch, setFBranch] = useState<"ankara" | "istanbul">("ankara");
  const [fCategory, setFCategory] = useState("muhtelif");
  const [fDesc, setFDesc] = useState("");
  const [fAmount, setFAmount] = useState("");
  const [fMethod, setFMethod] = useState<GiderYontem>("nakit");
  const [fSupplier, setFSupplier] = useState("");
  const [saving, setSaving] = useState(false);

  const load = useCallback(() => {
    setLoading(true);
    fetch(`/api/finans/gider?ay=${ay}${sube ? `&sube=${sube}` : ""}`)
      .then((r) => r.json())
      .then((d) => {
        if (d.ok) {
          setRecords(d.records || []);
          setErr("");
        } else setErr(d.error || "Yüklenemedi");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [ay, sube]);

  useEffect(() => {
    load();
  }, [load]);

  const toplam = useMemo(
    () => records.filter((g) => g.currency === "TL").reduce((s, g) => s + g.amount, 0),
    [records]
  );
  const kategoriler = useMemo(() => {
    const m = new Map<string, number>();
    for (const g of records) {
      if (g.currency !== "TL") continue;
      m.set(g.category, (m.get(g.category) || 0) + g.amount);
    }
    return [...m.entries()].sort((a, b) => b[1] - a[1]);
  }, [records]);

  async function kaydet() {
    const tutar = parseFloat(fAmount.replace(",", ".")) || 0;
    if (tutar <= 0 || !fCategory.trim()) return;
    setSaving(true);
    try {
      const r = await fetch("/api/finans/gider", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dateKey: fDate,
          branch: fBranch,
          category: fCategory,
          description: fDesc,
          amount: tutar,
          method: fMethod,
          supplier: fSupplier || undefined,
        }),
      });
      const d = await r.json();
      if (!d.ok) throw new Error(d.error);
      setFDesc("");
      setFAmount("");
      setFSupplier("");
      setFormOpen(false);
      load();
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Kaydedilemedi");
    } finally {
      setSaving(false);
    }
  }

  async function sil(g: Gider) {
    if (!confirm(`${g.category} — ₺${fmt(g.amount)} gideri silinsin mi?`)) return;
    await fetch(`/api/finans/gider?id=${g.id}&ay=${g.dateKey.slice(0, 7)}`, {
      method: "DELETE",
    });
    load();
  }

  return (
    <div>
      <div className="card no-print" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <input type="month" style={{ width: "auto" }} value={ay} onChange={(e) => setAy(e.target.value)} />
        <select style={{ width: "auto" }} value={sube} onChange={(e) => setSube(e.target.value)}>
          <option value="">Tüm Şubeler</option>
          <option value="ankara">Ankara</option>
          <option value="istanbul">İstanbul</option>
        </select>
        <span style={{ flex: 1 }} />
        <strong style={{ color: "var(--error)" }}>Toplam: ₺{fmt(toplam)}</strong>
        <button className="btn small" onClick={() => setFormOpen((o) => !o)}>
          {formOpen ? "Vazgeç" : "+ Gider Ekle"}
        </button>
      </div>

      {formOpen && (
        <div className="card">
          <div className="rw-grid2">
            <div>
              <label>Tarih</label>
              <input type="date" value={fDate} onChange={(e) => setFDate(e.target.value)} />
            </div>
            <div>
              <label>Şube</label>
              <select value={fBranch} onChange={(e) => setFBranch(e.target.value as "ankara" | "istanbul")}>
                <option value="ankara">Ankara</option>
                <option value="istanbul">İstanbul</option>
              </select>
            </div>
            <div>
              <label>Kategori</label>
              <input
                list="gider-kategoriler"
                value={fCategory}
                onChange={(e) => setFCategory(e.target.value)}
              />
              <datalist id="gider-kategoriler">
                {GIDER_KATEGORILERI.map((k) => (
                  <option key={k} value={k} />
                ))}
              </datalist>
            </div>
            <div>
              <label>Tutar (₺)</label>
              <input type="number" step="0.01" min="0" value={fAmount} onChange={(e) => setFAmount(e.target.value)} />
            </div>
            <div>
              <label>Yöntem</label>
              <select value={fMethod} onChange={(e) => setFMethod(e.target.value as GiderYontem)}>
                {Object.entries(GIDER_YONTEM_LABELS).map(([k, v]) => (
                  <option key={k} value={k}>{v}</option>
                ))}
              </select>
            </div>
            <div>
              <label>Ödenen Taraf (opsiyonel)</label>
              <input value={fSupplier} onChange={(e) => setFSupplier(e.target.value)} placeholder="tedarikçi / kurum" />
            </div>
            <div style={{ gridColumn: "1 / -1" }}>
              <label>Açıklama</label>
              <input value={fDesc} onChange={(e) => setFDesc(e.target.value)} placeholder="örn. Temmuz kirası" />
            </div>
          </div>
          <div style={{ marginTop: 12 }}>
            <button className="btn" onClick={kaydet} disabled={saving}>
              {saving ? "Kaydediliyor…" : "Kaydet"}
            </button>
          </div>
        </div>
      )}

      {err && <div className="notice err">{err}</div>}

      {kategoriler.length > 0 && (
        <div className="card" style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
          {kategoriler.map(([k, v]) => (
            <span key={k} style={{ fontSize: 13.5 }}>
              <strong>{k}</strong>: ₺{fmt(v)}
            </span>
          ))}
        </div>
      )}

      {loading ? (
        <p style={{ color: "var(--muted)" }}>Yükleniyor…</p>
      ) : (
        <div className="card" style={{ padding: 0, overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Tarih</th>
                <th>Şube</th>
                <th>Kategori</th>
                <th>Açıklama</th>
                <th>Yöntem</th>
                <th style={{ textAlign: "right" }}>Tutar</th>
                <th>Kaydeden</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {records.map((g) => (
                <tr key={g.id}>
                  <td>{g.dateKey.split("-").reverse().join(".")}</td>
                  <td style={{ fontSize: 12.5 }}>{g.branch === "istanbul" ? "İST" : "ANK"}</td>
                  <td style={{ fontWeight: 600 }}>{g.category}</td>
                  <td style={{ fontSize: 13 }}>{[g.description, g.supplier].filter(Boolean).join(" — ")}</td>
                  <td style={{ fontSize: 12.5 }}>{GIDER_YONTEM_LABELS[g.method]}</td>
                  <td style={{ textAlign: "right", color: "var(--error)", fontWeight: 600 }}>
                    {g.currency === "TL" ? "₺" : g.currency === "USD" ? "$" : "€"}{fmt(g.amount)}
                  </td>
                  <td style={{ fontSize: 12.5 }}>{g.createdBy}</td>
                  <td>
                    <button className="btn small danger" onClick={() => sil(g)}>🗑</button>
                  </td>
                </tr>
              ))}
              {!records.length && (
                <tr>
                  <td colSpan={8} style={{ color: "var(--muted)" }}>Bu ayda gider kaydı yok.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
