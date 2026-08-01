"use client";

import { useState } from "react";

export default function StockUpload() {
  const [file, setFile] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<{ ok: boolean; msg: string } | null>(null);

  async function upload() {
    if (!file) {
      setResult({ ok: false, msg: "Önce dosya seçin." });
      return;
    }
    setBusy(true);
    setResult(null);
    try {
      const form = new FormData();
      form.append("file", file);
      const res = await fetch("/api/stock", { method: "POST", body: form });
      const data = await res.json();
      if (data.ok) {
        setResult({
          ok: true,
          msg: `✅ Stok güncellendi: ${data.count} profil kalemi — Ankara ${data.ankaraTotal.toLocaleString("tr-TR")} mt, İstanbul ${data.istanbulTotal.toLocaleString("tr-TR")} mt. Site anında güncel veriyi gösteriyor.`,
        });
        setFile(null);
      } else {
        setResult({ ok: false, msg: data.error || "Yükleme başarısız." });
      }
    } catch {
      setResult({ ok: false, msg: "Sunucuya ulaşılamadı." });
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="card">
      <p style={{ color: "var(--text-2)", marginBottom: 16 }}>
        Muhasebe programından aldığınız günlük stok Excel&apos;ini (xls/xlsx)
        buraya yükleyin. Sistem <strong>DEPO ADI / STOK İSMİ / MİKTAR</strong>{" "}
        kolonlarını otomatik bulur, yalnızca çerçeve profillerini alır ve
        metrajı 2,9&apos;a bölerek <strong>boy</strong> olarak yayınlar.
      </p>
      <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
        <input
          type="file"
          accept=".xls,.xlsx"
          style={{ maxWidth: 380 }}
          onChange={(e) => setFile(e.target.files?.[0] || null)}
        />
        <button className="btn" onClick={upload} disabled={busy || !file}>
          {busy ? "Yükleniyor…" : "Stokları Güncelle"}
        </button>
      </div>
      {result && (
        <div className={`notice ${result.ok ? "ok" : "err"}`} style={{ marginTop: 14 }}>
          {result.msg}
        </div>
      )}
    </div>
  );
}
