"use client";

// Finans genel bakış — kasa özeti, aylık tahsilat/gider/kâr çubukları,
// vadesi yaklaşan çekler. Yalnızca aylık özet dosyalarını okur (hızlı).

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import type { FinansOzet, SubeKey } from "@/lib/finans-ozet";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
const fmt0 = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });

const AY_ADI = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];

interface VadeSatir {
  id: string;
  tur: string;
  kind: string;
  vade: string;
  tutar: number;
  kimden: string;
  branch: string;
  gecmis: boolean;
}

interface PortfoyOzet {
  alinanAdet: number;
  alinanToplam: number;
  verilenAdet: number;
  verilenToplam: number;
}

export default function FinansDashboard() {
  const [months, setMonths] = useState<string[]>([]);
  const [ozetler, setOzetler] = useState<FinansOzet[]>([]);
  const [vadeler, setVadeler] = useState<VadeSatir[]>([]);
  const [portfoy, setPortfoy] = useState<PortfoyOzet | null>(null);
  const [sube, setSube] = useState<"" | "ankara" | "istanbul">("");
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");

  useEffect(() => {
    fetch("/api/finans/ozet")
      .then((r) => r.json())
      .then((d) => {
        if (d.ok) {
          setMonths(d.months || []);
          setOzetler(d.ozetler || []);
          setVadeler(d.vadesiYaklasan || []);
          setPortfoy(d.portfoyOzet || null);
        } else setErr(d.error || "Yüklenemedi");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, []);

  // Ay → {tahsilat, gider, kar} — şube filtresine göre
  const seriler = useMemo(() => {
    const map = new Map(ozetler.map((o) => [o.month, o]));
    return months.map((m) => {
      const o = map.get(m);
      let tahsilat = 0, gider = 0;
      if (o) {
        const keys: SubeKey[] = sube ? [sube] : ["ankara", "istanbul", "belirsiz"];
        for (const k of keys) {
          const s = o.sube[k];
          if (!s) continue;
          tahsilat += s.tahsilatToplam;
          gider += s.giderToplam;
        }
      }
      return { month: m, tahsilat, gider, kar: tahsilat - gider };
    });
  }, [months, ozetler, sube]);

  const buAy = seriler[seriler.length - 1];
  const maxDeger = Math.max(1, ...seriler.map((s) => Math.max(s.tahsilat, s.gider)));

  if (loading) return <p style={{ color: "var(--muted)" }}>Yükleniyor…</p>;
  if (err) return <div className="notice err">{err}</div>;

  return (
    <div>
      <div className="no-print" style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        {(["", "ankara", "istanbul"] as const).map((s) => (
          <button
            key={s || "tum"}
            className={`btn small ${sube === s ? "" : "secondary"}`}
            onClick={() => setSube(s)}
          >
            {s === "" ? "Tümü" : s === "ankara" ? "Ankara" : "İstanbul"}
          </button>
        ))}
      </div>

      {/* Bu ay kartları */}
      <div className="cari-cards">
        <div className="cari-card">
          <span>Bu Ay Tahsilat</span>
          <strong style={{ color: "var(--success)" }}>₺{fmt(buAy?.tahsilat || 0)}</strong>
        </div>
        <div className="cari-card">
          <span>Bu Ay Gider</span>
          <strong style={{ color: "var(--error)" }}>₺{fmt(buAy?.gider || 0)}</strong>
        </div>
        <div className={`cari-card ${(buAy?.kar || 0) < 0 ? "borc" : ""}`}>
          <span>Bu Ay Kasa Kârı</span>
          <strong style={{ color: (buAy?.kar || 0) >= 0 ? "var(--success)" : "var(--error)" }}>
            ₺{fmt(buAy?.kar || 0)}
          </strong>
        </div>
        {portfoy && (
          <div className="cari-card">
            <span>Çek/Senet Portföyü</span>
            <strong>₺{fmt(portfoy.alinanToplam)}</strong>
            <span style={{ fontSize: 12 }}>{portfoy.alinanAdet} alınan kayıt</span>
          </div>
        )}
      </div>

      {/* Aylık çubuklar */}
      <h2>Son 12 Ay — Tahsilat / Gider</h2>
      <div className="card" style={{ overflowX: "auto" }}>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-end", minWidth: 700, height: 190, padding: "8px 4px" }}>
          {seriler.map((s) => {
            const [yy, mm] = s.month.split("-");
            return (
              <div key={s.month} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                <div style={{ display: "flex", gap: 3, alignItems: "flex-end", height: 130 }}>
                  <div
                    title={`Tahsilat ₺${fmt(s.tahsilat)}`}
                    style={{ width: 16, height: Math.max(2, (s.tahsilat / maxDeger) * 130),
                      background: "var(--success, #067a55)", borderRadius: "3px 3px 0 0" }}
                  />
                  <div
                    title={`Gider ₺${fmt(s.gider)}`}
                    style={{ width: 16, height: Math.max(2, (s.gider / maxDeger) * 130),
                      background: "var(--error, #b91c1c)", borderRadius: "3px 3px 0 0", opacity: 0.85 }}
                  />
                </div>
                <span style={{ fontSize: 11.5, color: "var(--muted)" }}>
                  {AY_ADI[Number(mm) - 1]} {yy.slice(2)}
                </span>
                <span style={{ fontSize: 11, fontWeight: 600,
                  color: s.kar >= 0 ? "var(--success)" : "var(--error)" }}>
                  {s.tahsilat || s.gider ? `${s.kar >= 0 ? "+" : ""}${fmt0(s.kar / 1000)}k` : ""}
                </span>
              </div>
            );
          })}
        </div>
        <p style={{ margin: "4px 0 0", fontSize: 12.5, color: "var(--muted)" }}>
          Yeşil: tahsilat · Kırmızı: gider · Alt satır: kasa kârı (bin ₺). Kâr,
          kasa bazlıdır (tahsilat − gider); çek portföyü tahsil edildikçe eklenir.
        </p>
      </div>

      {/* Vadesi yaklaşan çekler */}
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <h2 style={{ flex: 1 }}>Vadesi Yaklaşan Çek/Senet (30 gün)</h2>
        <Link href="/panel/finans/ceksenet" className="btn small secondary no-print">
          Portföye Git →
        </Link>
      </div>
      {vadeler.length === 0 ? (
        <div className="card" style={{ color: "var(--muted)", textAlign: "center" }}>
          Önümüzdeki 30 günde vadesi dolan kayıt yok.
        </div>
      ) : (
        <div className="card" style={{ padding: 0, overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Vade</th>
                <th>Tür</th>
                <th>Kimden / Kime</th>
                <th>Şube</th>
                <th style={{ textAlign: "right" }}>Tutar</th>
              </tr>
            </thead>
            <tbody>
              {vadeler.map((v) => (
                <tr key={v.id}>
                  <td style={{ fontWeight: 600, color: v.gecmis ? "var(--error)" : undefined }}>
                    {v.vade.split("-").reverse().join(".")} {v.gecmis && "⚠"}
                  </td>
                  <td style={{ fontSize: 12.5 }}>
                    {v.tur === "alinan" ? "Alınan" : "Verilen"} {v.kind === "cek" ? "çek" : "senet"}
                  </td>
                  <td>{v.kimden}</td>
                  <td style={{ fontSize: 12.5 }}>{v.branch === "istanbul" ? "İST" : "ANK"}</td>
                  <td style={{ textAlign: "right", fontWeight: 700 }}>₺{fmt(v.tutar)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
