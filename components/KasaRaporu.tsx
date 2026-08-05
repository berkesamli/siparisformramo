"use client";

// Kasa raporu — Excel'deki aylık kasa günlüğünün karşılığı.
// Tarih aralığı + şube filtresi; nakit/banka/döviz/portföy kırılımı.

import { useCallback, useEffect, useState } from "react";
import TahsilatModal from "./TahsilatModal";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

interface Satir {
  dateKey: string;
  yon: "G" | "C";
  tip: string;
  taraf: string;
  aciklama: string;
  kanal: string;
  currency: string;
  amount: number;
  branch: string;
  kaydeden: string;
}

interface Ozet {
  girisNakit: number;
  girisBanka: number;
  cikisNakit: number;
  cikisBanka: number;
  portfoyGiris: number;
  dovizUsd: number;
  dovizEur: number;
}

const bugun = () =>
  new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" });

export default function KasaRaporu() {
  const [bas, setBas] = useState(bugun().slice(0, 7) + "-01");
  const [son, setSon] = useState(bugun());
  const [sube, setSube] = useState("");
  const [rows, setRows] = useState<Satir[]>([]);
  const [ozet, setOzet] = useState<Ozet | null>(null);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [eldenAcik, setEldenAcik] = useState(false);

  const load = useCallback(() => {
    setLoading(true);
    fetch(`/api/finans/kasa?bas=${bas}&son=${son}${sube ? `&sube=${sube}` : ""}`)
      .then((r) => r.json())
      .then((d) => {
        if (d.ok) {
          setRows(d.rows || []);
          setOzet(d.ozet || null);
          setErr("");
        } else setErr(d.error || "Yüklenemedi");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [bas, son, sube]);

  useEffect(() => {
    load();
  }, [load]);

  const net =
    (ozet?.girisNakit || 0) + (ozet?.girisBanka || 0) -
    (ozet?.cikisNakit || 0) - (ozet?.cikisBanka || 0);

  return (
    <div>
      <div className="card no-print" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <button
          className={`btn small ${bas === bugun().slice(0, 7) + "-01" && son === bugun() ? "" : "secondary"}`}
          onClick={() => { setBas(bugun().slice(0, 7) + "-01"); setSon(bugun()); }}
        >
          Bu Ay
        </button>
        <button
          className={`btn small ${bas === bugun().slice(0, 4) + "-01-01" ? "" : "secondary"}`}
          onClick={() => { setBas(bugun().slice(0, 4) + "-01-01"); setSon(bugun()); }}
        >
          {bugun().slice(0, 4)} Tümü
        </button>
        <input type="date" style={{ width: "auto" }} value={bas} onChange={(e) => setBas(e.target.value)} />
        <span style={{ color: "var(--muted)" }}>→</span>
        <input type="date" style={{ width: "auto" }} value={son} onChange={(e) => setSon(e.target.value)} />
        <select style={{ width: "auto" }} value={sube} onChange={(e) => setSube(e.target.value)}>
          <option value="">Tüm Şubeler</option>
          <option value="ankara">Ankara</option>
          <option value="istanbul">İstanbul</option>
        </select>
        <span style={{ flex: 1 }} />
        <span style={{ fontSize: 13, color: "var(--muted)" }}>{rows.length} hareket</span>
        <button className="btn small" onClick={() => setEldenAcik(true)}>
          + Elden Tahsilat
        </button>
      </div>

      {eldenAcik && (
        <TahsilatModal
          baglam={{ customerName: "PERAKENDE", serbest: true, branch: (sube as "ankara" | "istanbul") || "ankara" }}
          onClose={() => setEldenAcik(false)}
          onSaved={load}
        />
      )}

      {ozet && (
        <div className="cari-cards">
          <div className="cari-card">
            <span>Giriş — Nakit</span>
            <strong style={{ color: "var(--success)" }}>₺{fmt(ozet.girisNakit)}</strong>
          </div>
          <div className="cari-card">
            <span>Giriş — Banka</span>
            <strong style={{ color: "var(--success)" }}>₺{fmt(ozet.girisBanka)}</strong>
            <span style={{ fontSize: 11.5 }}>havale + k.kartı + çek tahsili</span>
          </div>
          <div className="cari-card">
            <span>Çıkış — Nakit</span>
            <strong style={{ color: "var(--error)" }}>₺{fmt(ozet.cikisNakit)}</strong>
          </div>
          <div className="cari-card">
            <span>Çıkış — Banka</span>
            <strong style={{ color: "var(--error)" }}>₺{fmt(ozet.cikisBanka)}</strong>
          </div>
          <div className={`cari-card ${net < 0 ? "borc" : ""}`}>
            <span>Net Kasa Hareketi</span>
            <strong style={{ color: net >= 0 ? "var(--success)" : "var(--error)" }}>₺{fmt(net)}</strong>
          </div>
          {(ozet.portfoyGiris > 0 || ozet.dovizUsd !== 0 || ozet.dovizEur !== 0) && (
            <div className="cari-card">
              <span>Kasa Dışı</span>
              <strong style={{ fontSize: 15 }}>
                {ozet.portfoyGiris > 0 && <>çek/senet ₺{fmt(ozet.portfoyGiris)}</>}
              </strong>
              <span style={{ fontSize: 11.5 }}>
                {ozet.dovizUsd !== 0 && <>USD {fmt(ozet.dovizUsd)} · </>}
                {ozet.dovizEur !== 0 && <>EUR {fmt(ozet.dovizEur)}</>}
              </span>
            </div>
          )}
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
                <th>Tarih</th>
                <th>G/Ç</th>
                <th>Taraf</th>
                <th>Açıklama</th>
                <th>Kanal</th>
                <th>Şube</th>
                <th style={{ textAlign: "right" }}>Tutar</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i}>
                  <td style={{ whiteSpace: "nowrap" }}>{r.dateKey.split("-").reverse().join(".")}</td>
                  <td>
                    <span style={{ fontWeight: 700, color: r.yon === "G" ? "var(--success)" : "var(--error)" }}>
                      {r.yon === "G" ? "▲ G" : "▼ Ç"}
                    </span>
                  </td>
                  <td style={{ fontWeight: 600 }}>{r.taraf}</td>
                  <td style={{ fontSize: 12.5, maxWidth: 340 }}>{r.aciklama}</td>
                  <td style={{ fontSize: 12.5 }}>
                    {r.kanal === "portfoy" ? "çek/senet" : r.kanal}
                  </td>
                  <td style={{ fontSize: 12.5 }}>{r.branch === "istanbul" ? "İST" : "ANK"}</td>
                  <td style={{ textAlign: "right", fontWeight: 600,
                    color: r.yon === "G" ? "var(--success)" : "var(--error)" }}>
                    {r.currency === "TL" ? "₺" : r.currency === "USD" ? "$" : "€"}{fmt(r.amount)}
                  </td>
                </tr>
              ))}
              {!rows.length && (
                <tr>
                  <td colSpan={7} style={{ color: "var(--muted)" }}>Bu aralıkta hareket yok.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
