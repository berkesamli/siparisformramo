"use client";

// Çek/senet portföyü — vade sıralı liste, durum geçişleri, ciro takibi.
// "Çek kimde?" sorusunun cevabı durum rozetinden okunur.

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  CEKSENET_DURUM_LABELS,
  allowedTransitions,
  type CekSenet,
  type CekSenetDurum,
} from "@/lib/ceksenet";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const bugun = () =>
  new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" });

function vadeDurumu(vade: string): "gecti" | "yakin" | "normal" {
  const b = bugun();
  if (vade < b) return "gecti";
  const yedi = new Date(Date.now() + 7 * 86400000).toLocaleDateString("en-CA", {
    timeZone: "Europe/Istanbul",
  });
  return vade <= yedi ? "yakin" : "normal";
}

export default function CekSenetManager() {
  const [tur, setTur] = useState<"alinan" | "verilen">("alinan");
  const [durumFiltre, setDurumFiltre] = useState("portfoyde");
  const [records, setRecords] = useState<CekSenet[]>([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [formOpen, setFormOpen] = useState(false);

  // form
  const [f, setF] = useState({
    kind: "cek",
    branch: "ankara",
    customerName: "",
    supplier: "",
    cekSahibi: "",
    banka: "",
    belgeNo: "",
    tutar: "",
    vade: "",
    note: "",
  });
  const [saving, setSaving] = useState(false);

  const load = useCallback(() => {
    setLoading(true);
    fetch(`/api/finans/ceksenet?tur=${tur}${durumFiltre ? `&durum=${durumFiltre}` : ""}`)
      .then((r) => r.json())
      .then((d) => {
        if (d.ok) {
          setRecords(d.records || []);
          setErr("");
        } else setErr(d.error || "Yüklenemedi");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [tur, durumFiltre]);

  useEffect(() => {
    load();
  }, [load]);

  const toplam = useMemo(() => records.reduce((s, c) => s + c.tutar, 0), [records]);

  async function kaydet() {
    const tutar = parseFloat(f.tutar.replace(",", ".")) || 0;
    if (tutar <= 0 || !f.vade) return;
    setSaving(true);
    try {
      const r = await fetch("/api/finans/ceksenet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...f, tur, tutar }),
      });
      const d = await r.json();
      if (!d.ok) throw new Error(d.error);
      setF({ ...f, customerName: "", supplier: "", cekSahibi: "", banka: "", belgeNo: "", tutar: "", vade: "", note: "" });
      setFormOpen(false);
      load();
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Kaydedilemedi");
    } finally {
      setSaving(false);
    }
  }

  async function gecis(cs: CekSenet, hedef: CekSenetDurum) {
    let ciroTarget: string | undefined;
    if (hedef === "ciro") {
      ciroTarget = prompt("Hangi tedarikçiye ciro edildi?") || undefined;
      if (!ciroTarget) return;
    }
    const etiket = CEKSENET_DURUM_LABELS[hedef];
    if (!confirm(`${cs.customerName || cs.supplier} — ₺${fmt(cs.tutar)} → "${etiket}" olarak işaretlensin mi?`)) return;
    const r = await fetch("/api/finans/ceksenet", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: cs.id, durum: hedef, ciroTarget }),
    });
    const d = await r.json();
    if (!d.ok) setErr(d.error || "Güncellenemedi");
    load();
  }

  return (
    <div>
      <div className="card no-print" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <button className={`btn small ${tur === "alinan" ? "" : "secondary"}`} onClick={() => setTur("alinan")}>
          Alınan (Müşteri)
        </button>
        <button className={`btn small ${tur === "verilen" ? "" : "secondary"}`} onClick={() => setTur("verilen")}>
          Verilen (Tedarikçi)
        </button>
        <select style={{ width: "auto" }} value={durumFiltre} onChange={(e) => setDurumFiltre(e.target.value)}>
          <option value="portfoyde">Portföyde</option>
          <option value="">Tümü</option>
          {Object.entries(CEKSENET_DURUM_LABELS)
            .filter(([k]) => k !== "portfoyde")
            .map(([k, v]) => (
              <option key={k} value={k}>{v}</option>
            ))}
        </select>
        <span style={{ flex: 1 }} />
        <strong>
          {records.length} kayıt · ₺{fmt(toplam)}
        </strong>
        <button className="btn small" onClick={() => setFormOpen((o) => !o)}>
          {formOpen ? "Vazgeç" : "+ Yeni Kayıt"}
        </button>
      </div>

      {formOpen && (
        <div className="card">
          <div className="rw-grid2">
            <div>
              <label>Tür</label>
              <select value={f.kind} onChange={(e) => setF({ ...f, kind: e.target.value })}>
                <option value="cek">Çek</option>
                <option value="senet">Senet</option>
              </select>
            </div>
            <div>
              <label>Şube</label>
              <select value={f.branch} onChange={(e) => setF({ ...f, branch: e.target.value })}>
                <option value="ankara">Ankara</option>
                <option value="istanbul">İstanbul</option>
              </select>
            </div>
            {tur === "alinan" ? (
              <>
                <div>
                  <label>Müşteri (kimden alındı)</label>
                  <input value={f.customerName} onChange={(e) => setF({ ...f, customerName: e.target.value })} />
                </div>
                <div>
                  <label>Çek Sahibi (keşideci, farklıysa)</label>
                  <input value={f.cekSahibi} onChange={(e) => setF({ ...f, cekSahibi: e.target.value })} placeholder="üçüncü şahıs çeki ise" />
                </div>
              </>
            ) : (
              <div style={{ gridColumn: "1 / -1" }}>
                <label>Tedarikçi (kime verildi)</label>
                <input value={f.supplier} onChange={(e) => setF({ ...f, supplier: e.target.value })} />
              </div>
            )}
            <div>
              <label>Banka</label>
              <input value={f.banka} onChange={(e) => setF({ ...f, banka: e.target.value })} placeholder="senette boş bırakın" />
            </div>
            <div>
              <label>Çek/Senet No</label>
              <input value={f.belgeNo} onChange={(e) => setF({ ...f, belgeNo: e.target.value })} />
            </div>
            <div>
              <label>Tutar (₺)</label>
              <input type="number" step="0.01" min="0" value={f.tutar} onChange={(e) => setF({ ...f, tutar: e.target.value })} />
            </div>
            <div>
              <label>Vade</label>
              <input type="date" value={f.vade} onChange={(e) => setF({ ...f, vade: e.target.value })} />
            </div>
            <div style={{ gridColumn: "1 / -1" }}>
              <label>Not</label>
              <input value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} />
            </div>
          </div>
          {tur === "alinan" && (
            <p className="notice info" style={{ marginTop: 10 }}>
              Alınan çek/senet müşterinin carisini <strong>hemen düşürür</strong>;
              kasaya ise tahsil edildiğinde girer.
            </p>
          )}
          <div style={{ marginTop: 12 }}>
            <button className="btn" onClick={kaydet} disabled={saving}>
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
                <th>Vade</th>
                <th>Tür</th>
                <th>{tur === "alinan" ? "Müşteri" : "Tedarikçi"}</th>
                <th>Banka / No</th>
                <th style={{ textAlign: "right" }}>Tutar</th>
                <th>Durum</th>
                <th>Şube</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {records.map((cs) => {
                const vd = cs.durum === "portfoyde" ? vadeDurumu(cs.vade) : "normal";
                return (
                  <tr key={cs.id}>
                    <td style={{ whiteSpace: "nowrap", fontWeight: 600,
                      color: vd === "gecti" ? "var(--error)" : vd === "yakin" ? "#b45309" : undefined }}>
                      {cs.vade.split("-").reverse().join(".")}
                      {vd === "gecti" && " ⚠"}
                      {vd === "yakin" && " ⏰"}
                    </td>
                    <td style={{ fontSize: 12.5 }}>{cs.kind === "cek" ? "Çek" : "Senet"}</td>
                    <td>
                      {cs.customerName || cs.supplier}
                      {cs.cekSahibi && cs.cekSahibi !== cs.customerName && (
                        <div style={{ fontSize: 11.5, color: "var(--muted)" }}>keşideci: {cs.cekSahibi}</div>
                      )}
                    </td>
                    <td style={{ fontSize: 12.5 }}>{[cs.banka, cs.belgeNo].filter(Boolean).join(" / ") || "—"}</td>
                    <td style={{ textAlign: "right", fontWeight: 700 }}>₺{fmt(cs.tutar)}</td>
                    <td>
                      <span className={`badge ${cs.durum === "portfoyde" ? "var" : cs.durum === "karsiliksiz" ? "yok" : "az"}`}>
                        {CEKSENET_DURUM_LABELS[cs.durum]}
                      </span>
                      {cs.durum === "ciro" && cs.ciroTarget && (
                        <div style={{ fontSize: 11.5, color: "var(--muted)" }}>→ {cs.ciroTarget}</div>
                      )}
                    </td>
                    <td style={{ fontSize: 12.5 }}>{cs.branch === "istanbul" ? "İST" : "ANK"}</td>
                    <td style={{ whiteSpace: "nowrap" }}>
                      {allowedTransitions(cs.tur, cs.durum).map((h) => (
                        <button key={h} className="btn small secondary" style={{ marginRight: 4 }}
                          onClick={() => gecis(cs, h)}>
                          {CEKSENET_DURUM_LABELS[h]}
                        </button>
                      ))}
                    </td>
                  </tr>
                );
              })}
              {!records.length && (
                <tr>
                  <td colSpan={8} style={{ color: "var(--muted)" }}>Bu filtreye uyan kayıt yok.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
