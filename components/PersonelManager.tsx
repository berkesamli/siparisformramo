"use client";

// Personel avans/maaş takibi — Avans-Maaş Excel'inin karşılığı.
// Ödemeler gider olarak kaydedilir (kategori: maaş/avans/prim + personelId);
// bu ekran o giderleri kişi bazında toplayıp "maaş / çektiği / kalan" gösterir.

import { useCallback, useEffect, useMemo, useState } from "react";
import type { Personel } from "@/lib/personel";
import type { Gider } from "@/lib/gider";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const buAy = () =>
  new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" }).slice(0, 7);

export default function PersonelManager() {
  const [ay, setAy] = useState(buAy());
  const [personel, setPersonel] = useState<Personel[]>([]);
  const [odemeler, setOdemeler] = useState<Gider[]>([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");

  const [formOpen, setFormOpen] = useState(false);
  const [fName, setFName] = useState("");
  const [fBranch, setFBranch] = useState<"ankara" | "istanbul">("ankara");
  const [fStart, setFStart] = useState("");
  const [fSalary, setFSalary] = useState("");

  const [odemePersonel, setOdemePersonel] = useState<Personel | null>(null);
  const [oKategori, setOKategori] = useState("avans");
  const [oTutar, setOTutar] = useState("");
  const [oYontem, setOYontem] = useState("nakit");
  const [saving, setSaving] = useState(false);

  const load = useCallback(() => {
    setLoading(true);
    fetch(`/api/finans/personel?ay=${ay}`)
      .then((r) => r.json())
      .then((d) => {
        if (d.ok) {
          setPersonel(d.personel || []);
          setOdemeler(d.odemeler || []);
          setErr("");
        } else setErr(d.error || "Yüklenemedi");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [ay]);

  useEffect(() => {
    load();
  }, [load]);

  const kisiOdeme = useMemo(() => {
    const m = new Map<string, number>();
    for (const g of odemeler) {
      if (!g.personelId) continue;
      m.set(g.personelId, (m.get(g.personelId) || 0) + g.amount);
    }
    return m;
  }, [odemeler]);

  async function personelKaydet() {
    if (!fName.trim()) return;
    setSaving(true);
    try {
      const r = await fetch("/api/finans/personel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: fName,
          branch: fBranch,
          startDate: fStart || undefined,
          salary: parseFloat(fSalary.replace(",", ".")) || undefined,
        }),
      });
      const d = await r.json();
      if (!d.ok) throw new Error(d.error);
      setFName("");
      setFStart("");
      setFSalary("");
      setFormOpen(false);
      load();
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Kaydedilemedi");
    } finally {
      setSaving(false);
    }
  }

  async function odemeKaydet() {
    if (!odemePersonel) return;
    const tutar = parseFloat(oTutar.replace(",", ".")) || 0;
    if (tutar <= 0) return;
    setSaving(true);
    try {
      const r = await fetch("/api/finans/gider", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          branch: odemePersonel.branch,
          category: oKategori,
          description: `${odemePersonel.name} — ${oKategori}`,
          amount: tutar,
          method: oYontem,
          personelId: odemePersonel.id,
        }),
      });
      const d = await r.json();
      if (!d.ok) throw new Error(d.error);
      setOdemePersonel(null);
      setOTutar("");
      load();
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Kaydedilemedi");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div>
      <div className="card no-print" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <input type="month" style={{ width: "auto" }} value={ay} onChange={(e) => setAy(e.target.value)} />
        <span style={{ flex: 1 }} />
        <button className="btn small" onClick={() => setFormOpen((o) => !o)}>
          {formOpen ? "Vazgeç" : "+ Personel Ekle"}
        </button>
      </div>

      {formOpen && (
        <div className="card">
          <div className="rw-grid2">
            <div>
              <label>Ad Soyad</label>
              <input value={fName} onChange={(e) => setFName(e.target.value)} />
            </div>
            <div>
              <label>Şube</label>
              <select value={fBranch} onChange={(e) => setFBranch(e.target.value as "ankara" | "istanbul")}>
                <option value="ankara">Ankara</option>
                <option value="istanbul">İstanbul</option>
              </select>
            </div>
            <div>
              <label>İşe Başlama</label>
              <input type="date" value={fStart} onChange={(e) => setFStart(e.target.value)} />
            </div>
            <div>
              <label>Aylık Maaş (₺)</label>
              <input type="number" step="0.01" value={fSalary} onChange={(e) => setFSalary(e.target.value)} />
            </div>
          </div>
          <div style={{ marginTop: 12 }}>
            <button className="btn" onClick={personelKaydet} disabled={saving}>
              {saving ? "Kaydediliyor…" : "Kaydet"}
            </button>
          </div>
        </div>
      )}

      {err && <div className="notice err">{err}</div>}

      {loading ? (
        <p style={{ color: "var(--muted)" }}>Yükleniyor…</p>
      ) : (
        <div className="card" style={{ padding: 0, overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Ad Soyad</th>
                <th>Şube</th>
                <th>İşe Başlama</th>
                <th style={{ textAlign: "right" }}>Maaş</th>
                <th style={{ textAlign: "right" }}>Bu Ay Çektiği</th>
                <th style={{ textAlign: "right" }}>Kalan</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {personel.map((p) => {
                const cekti = kisiOdeme.get(p.id) || 0;
                const kalan = (p.salary || 0) - cekti;
                return (
                  <tr key={p.id} style={p.endDate ? { opacity: 0.55 } : undefined}>
                    <td style={{ fontWeight: 600 }}>
                      {p.name}
                      {p.endDate && <span style={{ fontSize: 11.5, color: "var(--muted)" }}> (ayrıldı)</span>}
                    </td>
                    <td style={{ fontSize: 12.5 }}>{p.branch === "istanbul" ? "İST" : "ANK"}</td>
                    <td style={{ fontSize: 12.5 }}>{p.startDate?.split("-").reverse().join(".") || "—"}</td>
                    <td style={{ textAlign: "right" }}>{p.salary ? `₺${fmt(p.salary)}` : "—"}</td>
                    <td style={{ textAlign: "right", color: "var(--error)" }}>₺{fmt(cekti)}</td>
                    <td style={{ textAlign: "right", fontWeight: 600, color: kalan < 0 ? "var(--error)" : "var(--success)" }}>
                      {p.salary ? `₺${fmt(kalan)}` : "—"}
                    </td>
                    <td>
                      <button className="btn small secondary" onClick={() => setOdemePersonel(p)}>
                        💸 Ödeme
                      </button>
                    </td>
                  </tr>
                );
              })}
              {!personel.length && (
                <tr>
                  <td colSpan={7} style={{ color: "var(--muted)" }}>
                    Henüz personel kartı yok. &quot;+ Personel Ekle&quot; ile başlayın.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      <h2 style={{ marginTop: 24 }}>Bu Ayın Ödemeleri</h2>
      <div className="card" style={{ padding: 0, overflowX: "auto" }}>
        <table>
          <thead>
            <tr>
              <th>Tarih</th>
              <th>Kişi / Açıklama</th>
              <th>Kategori</th>
              <th style={{ textAlign: "right" }}>Tutar</th>
            </tr>
          </thead>
          <tbody>
            {odemeler.map((g) => (
              <tr key={g.id}>
                <td>{g.dateKey.split("-").reverse().join(".")}</td>
                <td>{g.description}</td>
                <td style={{ fontSize: 12.5 }}>{g.category}</td>
                <td style={{ textAlign: "right", color: "var(--error)" }}>₺{fmt(g.amount)}</td>
              </tr>
            ))}
            {!odemeler.length && (
              <tr>
                <td colSpan={4} style={{ color: "var(--muted)" }}>Bu ayda ödeme yok.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {odemePersonel && (
        <div className="modal-backdrop" onClick={() => setOdemePersonel(null)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <h2 style={{ marginTop: 0 }}>💸 {odemePersonel.name} — Ödeme</h2>
            <div className="rw-grid2">
              <div>
                <label>Kategori</label>
                <select value={oKategori} onChange={(e) => setOKategori(e.target.value)}>
                  <option value="avans">Avans</option>
                  <option value="maaş">Maaş</option>
                  <option value="prim">Prim</option>
                </select>
              </div>
              <div>
                <label>Yöntem</label>
                <select value={oYontem} onChange={(e) => setOYontem(e.target.value)}>
                  <option value="nakit">Nakit</option>
                  <option value="havale">Banka</option>
                </select>
              </div>
              <div style={{ gridColumn: "1 / -1" }}>
                <label>Tutar (₺)</label>
                <input type="number" step="0.01" min="0" value={oTutar} onChange={(e) => setOTutar(e.target.value)} />
              </div>
            </div>
            <div style={{ marginTop: 12, display: "flex", gap: 10 }}>
              <button className="btn" onClick={odemeKaydet} disabled={saving}>
                {saving ? "Kaydediliyor…" : "Kaydet"}
              </button>
              <button className="btn secondary" onClick={() => setOdemePersonel(null)}>
                Vazgeç
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
