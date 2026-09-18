"use client";

// Stok sayfası: Mikro'dan çekme kartı. Yayındaki verinin kaynağı/zamanı,
// "Mikro'dan şimdi çek" (kaydeder) ve sahipler için "Önizle" (kaydetmez;
// depoları, yöntemi, birimi ve örnek satırları gösterir — Excel ile karşılaştırmak için).

import { useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Bilgi {
  depolar: { no: number; ad: string }[];
  ankaraDepolar: number[];
  istanbulDepolar: number[];
  yontem?: string;
  birimler: string[];
  hamSatir: number;
  profilSatir: number;
  kalem: number;
  sureMs: number;
  ornek: { kod: string; isim: string; birim: string; ankara: number; istanbul: number; cikarilanKod: string }[];
}
interface Sonuc {
  ok: boolean; error?: string; kaydedildi?: boolean; count?: number; ankaraTotal?: number; istanbulTotal?: number;
  updatedAt?: string; sourceName?: string; bilgi?: Bilgi;
}

const nf = (n: number) => (Number(n) || 0).toLocaleString("tr-TR");
const zaman = (iso: string) => new Date(iso).toLocaleString("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });

export default function StockMikro({ owner = false }: { owner?: boolean }) {
  const [yayin, setYayin] = useState<{ updatedAt: string; sourceName: string; kalem: number } | null>(null);
  const [mikroKurulu, setMikroKurulu] = useState<boolean | null>(null);
  const [calisiyor, setCalisiyor] = useState<"" | "cek" | "onizle">("");
  const [sonuc, setSonuc] = useState<Sonuc | null>(null);
  const [onizleme, setOnizleme] = useState(false);

  async function yayinOku() {
    try {
      const r = await fetch("/api/stock");
      const d = await r.json();
      if (d.ok && d.data) setYayin({ updatedAt: d.data.updatedAt, sourceName: d.data.sourceName, kalem: (d.data.items || []).length });
    } catch { /* kart yine çizilir */ }
  }
  useEffect(() => {
    yayinOku();
    fetch("/api/mikro/test").then((r) => r.json()).then((d) => setMikroKurulu(d.ok ? !!d.kurulu : null)).catch(() => setMikroKurulu(null));
  }, []);

  async function calistir(kaydet: boolean) {
    setCalisiyor(kaydet ? "cek" : "onizle"); setSonuc(null); setOnizleme(!kaydet);
    try {
      const r = await fetch("/api/stock/mikro", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ kaydet }) });
      const d = (await r.json()) as Sonuc;
      setSonuc(d);
      if (d.ok && d.kaydedildi) yayinOku();
    } catch { setSonuc({ ok: false, error: "Sunucuya ulaşılamadı." }); }
    finally { setCalisiyor(""); }
  }

  const mikroKaynak = yayin?.sourceName?.startsWith("Mikro");
  const b = sonuc?.bilgi;

  return (
    <div className="card">
      <div className="card-head">
        <span className="card-head-icon"><Icon name="refresh" size={18} /></span>
        <div>
          <h2>Mikro&apos;dan Stok</h2>
          <span className="card-head-sub">Excel yüklemesinin yerine: aynı süzgeç ve formüllerle doğrudan Mikro&apos;dan</span>
        </div>
        <span className="spacer" />
        <div className="card-head-actions">
          {owner && (
            <button type="button" className="btn small secondary" onClick={() => calistir(false)} disabled={!!calisiyor || mikroKurulu === false} title="Kaydetmeden çeker; depoları ve örnek satırları gösterir">
              <Icon name="eye" size={15} /> {calisiyor === "onizle" ? "Çekiliyor…" : "Önizle"}
            </button>
          )}
          <button type="button" className="btn small" onClick={() => calistir(true)} disabled={!!calisiyor || mikroKurulu === false}>
            <Icon name="download" size={15} /> {calisiyor === "cek" ? "Çekiliyor…" : "Mikro'dan şimdi çek"}
          </button>
        </div>
      </div>

      <div className="row" style={{ gap: 14, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div style={{ flex: "1 1 260px", minWidth: 0 }}>
          <div className="muted" style={{ fontSize: 11.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em" }}>Yayındaki stok</div>
          {yayin ? (
            <div style={{ marginTop: 4 }}>
              <strong>{zaman(yayin.updatedAt)}</strong> · {nf(yayin.kalem)} profil kalemi
              <div className="muted" style={{ fontSize: 12.5 }}>
                Kaynak: {yayin.sourceName} {mikroKaynak ? <span className="badge ok" style={{ marginLeft: 6 }}>Mikro</span> : <span className="badge" style={{ marginLeft: 6 }}>Excel</span>}
              </div>
            </div>
          ) : <div className="skeleton" style={{ height: 38, marginTop: 4 }} />}
        </div>
        <div style={{ flex: "1 1 260px", minWidth: 0 }} className="muted">
          <div style={{ fontSize: 12.5 }}>
            Çalışanlar stok sorguladığında veri 2 saatten eskiyse Mikro&apos;dan kendiliğinden tazelenir; ayrıca her sabah 07:30&apos;da çekilir. Excel yüklemesi yedek olarak kalır.
          </div>
          {mikroKurulu === false && <div className="notice warn" style={{ marginTop: 8 }}>Mikro bağlantısı ayarlanmamış (Bildirim Ayarları → Mikro Bağlantısı).</div>}
        </div>
      </div>

      {sonuc && (
        <div className={`notice ${sonuc.ok ? "ok" : "err"}`} style={{ marginTop: 14 }}>
          {sonuc.ok ? (
            <>
              <strong>{sonuc.kaydedildi ? "Stok Mikro'dan güncellendi." : "Önizleme (kaydedilmedi)."}</strong>{" "}
              {nf(sonuc.count || 0)} profil kalemi — Ankara {nf(sonuc.ankaraTotal || 0)} mt, İstanbul {nf(sonuc.istanbulTotal || 0)} mt
              {b && <> · {b.yontem} · {(b.sureMs / 1000).toFixed(1)} sn</>}
            </>
          ) : (
            <><strong>Çekilemedi.</strong> {sonuc.error}</>
          )}
        </div>
      )}

      {b && (onizleme || !sonuc?.ok) && (
        <details open style={{ marginTop: 10 }}>
          <summary className="muted" style={{ cursor: "pointer", fontSize: 13 }}>Ayrıntı: depolar, yöntem, örnek satırlar</summary>
          <div style={{ fontSize: 13, marginTop: 8 }}>
            <div><strong>Depolar:</strong> {b.depolar.length ? b.depolar.map((d) => `${d.no}: ${d.ad}`).join(" · ") : "okunamadı"}</div>
            <div><strong>Ankara'ya sayılan:</strong> {b.ankaraDepolar.join(", ") || "—"} · <strong>İstanbul'a sayılan:</strong> {b.istanbulDepolar.join(", ") || "—"}</div>
            <div><strong>Yöntem:</strong> {b.yontem || "—"} · <strong>Birim:</strong> {b.birimler.join(", ") || "—"} · {nf(b.hamSatir)} satır geldi, {nf(b.profilSatir)} PROFİL, {nf(b.kalem)} kalem</div>
            {b.ornek.length > 0 && (
              <div className="table-wrap" style={{ marginTop: 8 }}>
                <table>
                  <thead><tr><th>Mikro stok adı</th><th>Sistemdeki kod</th><th>Birim</th><th className="num">Ankara</th><th className="num">İstanbul</th></tr></thead>
                  <tbody>
                    {b.ornek.map((o) => (
                      <tr key={o.kod}>
                        <td style={{ fontSize: 12.5 }}>{o.isim}</td>
                        <td><code>{o.cikarilanKod}</code></td>
                        <td>{o.birim || "—"}</td>
                        <td className="num">{nf(o.ankara)}</td>
                        <td className="num">{nf(o.istanbul)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </details>
      )}
    </div>
  );
}
