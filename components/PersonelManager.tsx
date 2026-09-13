"use client";

// Personel avans/maaş takibi — Avans-Maaş Excel'inin karşılığı.
// Ödemeler gider olarak kaydedilir (kategori: maaş/avans/prim + personelId);
// bu ekran o giderleri kişi bazında toplayıp "maaş / çektiği / kalan" gösterir.

import { useCallback, useEffect, useMemo, useState } from "react";
import Icon from "@/components/shell/Icon";
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
        <input type="month" style={{ width: "auto", maxWidth: "100%" }} value={ay} onChange={(e) => setAy(e.target.value)} />
        <span className="spacer" />
        <button type="button" className="btn small" onClick={() => setFormOpen((o) => !o)}>
          {formOpen ? "Vazgeç" : "+ Personel Ekle"}
        </button>
      </div>

      {formOpen && (
        <div className="card no-print">
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
            <button type="button" className="btn" onClick={personelKaydet} disabled={saving}>
              {saving ? "Kaydediliyor…" : "Kaydet"}
            </button>
          </div>
        </div>
      )}

      {err && <div className="notice err">{err}</div>}

      {loading ? (
        <p className="muted">Yükleniyor…</p>
      ) : (
        <div className="card pad-0">
          <div className="table-wrap" style={{ margin: 0, padding: 0 }}>
            <table>
              <thead>
                <tr>
                  <th>Ad Soyad</th>
                  <th>Şube</th>
                  <th>İşe Başlama</th>
                  <th className="num">Maaş</th>
                  <th className="num">Bu Ay Çektiği</th>
                  <th className="num">Kalan</th>
                  <th className="no-print"></th>
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
                        {p.endDate && <span className="muted" style={{ fontSize: 11.5 }}> (ayrıldı)</span>}
                      </td>
                      <td style={{ fontSize: 12.5 }}>{p.branch === "istanbul" ? "İST" : "ANK"}</td>
                      <td style={{ fontSize: 12.5, whiteSpace: "nowrap" }}>{p.startDate?.split("-").reverse().join(".") || "—"}</td>
                      <td className="num" style={{ whiteSpace: "nowrap" }}>{p.salary ? `₺${fmt(p.salary)}` : "—"}</td>
                      <td className="num" style={{ color: "var(--error)", whiteSpace: "nowrap" }}>₺{fmt(cekti)}</td>
                      <td className="num" style={{ fontWeight: 600, whiteSpace: "nowrap", color: kalan < 0 ? "var(--error)" : "var(--success)" }}>
                        {p.salary ? `₺${fmt(kalan)}` : "—"}
                      </td>
                      <td className="no-print">
                        <button type="button" className="btn small secondary" onClick={() => setOdemePersonel(p)}>
                          <Icon name="wallet" size={14} /> Ödeme
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {!personel.length && (
                  <tr>
                    <td colSpan={7}>
                      <div className="empty" style={{ padding: "22px 12px" }}>
                        Henüz personel kartı yok. &quot;+ Personel Ekle&quot; ile başlayın.
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <h2 style={{ marginTop: 24 }}>Bu Ayın Ödemeleri</h2>
      <div className="card pad-0">
        <div className="table-wrap" style={{ margin: 0, padding: 0 }}>
          <table>
            <thead>
              <tr>
                <th>Tarih</th>
                <th>Kişi / Açıklama</th>
                <th>Kategori</th>
                <th className="num">Tutar</th>
              </tr>
            </thead>
            <tbody>
              {odemeler.map((g) => (
                <tr key={g.id}>
                  <td style={{ whiteSpace: "nowrap" }}>{g.dateKey.split("-").reverse().join(".")}</td>
                  <td>{g.description}</td>
                  <td style={{ fontSize: 12.5 }}>{g.category}</td>
                  <td className="num" style={{ color: "var(--error)", whiteSpace: "nowrap" }}>₺{fmt(g.amount)}</td>
                </tr>
              ))}
              {!odemeler.length && (
                <tr>
                  <td colSpan={4}>
                    <div className="empty" style={{ padding: "22px 12px" }}>Bu ayda ödeme yok.</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {odemePersonel && (
        <div className="modal-backdrop" onClick={() => setOdemePersonel(null)}>
          <div className="modal-card" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
            <h2 style={{ marginTop: 0, display: "flex", alignItems: "center", gap: 8 }}>
              <span className="card-head-icon" style={{ width: 30, height: 30 }}><Icon name="wallet" size={16} /></span>
              <span style={{ minWidth: 0, overflowWrap: "anywhere" }}>{odemePersonel.name} — Ödeme</span>
            </h2>
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
            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button type="button" className="btn" onClick={odemeKaydet} disabled={saving}>
                {saving ? "Kaydediliyor…" : "Kaydet"}
              </button>
              <button type="button" className="btn secondary" onClick={() => setOdemePersonel(null)}>
                Vazgeç
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
