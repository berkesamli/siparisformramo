"use client";

// Parti (konteyner) bazlı alış fiyatları + kod bazlı satış/kâr analizi.
// Yalnızca maliyet yetkisi olanlar görür (varsayılan: Berke + Özgür).

import { useCallback, useEffect, useMemo, useState } from "react";
import { FRAME_PROFILES } from "@/data/catalog";
import { GLASS_TYPES } from "@/data/glass";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import type { MaliyetData, Parti } from "@/lib/maliyet";
import { eslesir } from "@/lib/search-norm";

// Kalem türleri: çerçeve metre, cam-ayna m², teknik malzeme kutu bazlı.
const TURLER = [
  { key: "cerceve", ad: "Çerçeve", birim: "mt" as const },
  { key: "cam", ad: "Cam / Ayna", birim: "m2" as const },
  { key: "teknik", ad: "Teknik Malzeme", birim: "kutu" as const },
];
const CAM_SECENEKLERI = [...GLASS_TYPES.map((g) => g.name), "Ayna"];
const birimEtiket = (b?: string) => (b === "m2" ? "m²" : b === "kutu" ? "kutu" : "mt");

const fmt = (n: number, d = 2) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: d,
    maximumFractionDigits: d,
  });

const simge = (c: string) => (c === "USD" ? "$" : c === "EUR" ? "€" : "₺");

interface AnalizSatir {
  code: string;
  metraj: number;
  ciro: number;
  maliyet: number | null;
  kar: number | null;
  marj: number | null;
  durum: string;
  satirSayisi: number;
}

interface Ozet {
  toplamCiro: number;
  maliyetliCiro: number;
  toplamMaliyet: number;
  toplamKar: number;
  kapsam: number;
}

const bugun = () =>
  new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" });

export default function MaliyetManager() {
  const [tab, setTab] = useState<"fiyat" | "analiz">("fiyat");
  const [data, setData] = useState<MaliyetData | null>(null);
  const [seciliParti, setSeciliParti] = useState("");
  const [analiz, setAnaliz] = useState<AnalizSatir[]>([]);
  const [ozet, setOzet] = useState<Ozet | null>(null);
  const [ay, setAy] = useState("");
  const [ara, setAra] = useState("");
  const [err, setErr] = useState("");
  const [msg, setMsg] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  // yeni parti formu
  const [ypAcik, setYpAcik] = useState(false);
  const [ypAd, setYpAd] = useState("");
  const [ypTarih, setYpTarih] = useState(bugun());

  // kalem formu
  const [fTur, setFTur] = useState("cerceve");
  const [fKod, setFKod] = useState("");
  const [fAlis, setFAlis] = useState("");
  const [fBirim, setFBirim] = useState("USD");
  const [pctInput, setPctInput] = useState("");
  const turBilgi = TURLER.find((t) => t.key === fTur) || TURLER[0];

  const load = useCallback(() => {
    setLoading(true);
    const url =
      tab === "analiz"
        ? `/api/maliyet?analiz=1${ay ? `&ay=${ay}` : ""}`
        : "/api/maliyet";
    fetch(url)
      .then((r) => r.json())
      .then((d) => {
        if (!d.ok) { setErr(d.error || "Yüklenemedi"); return; }
        setData(d.data);
        if (d.analiz) { setAnaliz(d.analiz); setOzet(d.ozet); }
        setErr("");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [tab, ay]);

  useEffect(() => { load(); }, [load]);

  const partiler = useMemo(
    () => [...(data?.partiler || [])].sort((a, b) => b.tarih.localeCompare(a.tarih)),
    [data]
  );
  const parti: Parti | undefined =
    partiler.find((x) => x.id === seciliParti) || partiler[0];

  useEffect(() => {
    if (parti && parti.id !== seciliParti) setSeciliParti(parti.id);
    setPctInput(parti?.pct != null ? String(parti.pct) : "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [parti?.id]);

  async function gonder(body: object) {
    setSaving(true); setMsg("");
    try {
      const r = await fetch("/api/maliyet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const d = await r.json();
      if (!d.ok) throw new Error(d.error);
      setData(d.data);
      setMsg("Kaydedildi.");
      setTimeout(() => setMsg(""), 2000);
      return d.data as MaliyetData;
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Kaydedilemedi");
      return null;
    } finally {
      setSaving(false);
    }
  }

  const kalemler = useMemo(() => {
    const list = Object.values(parti?.items || {});
    list.sort((a, b) => a.code.localeCompare(b.code, "tr"));
    return ara.trim() ? list.filter((k) => eslesir(ara, k.code)) : list;
  }, [parti, ara]);

  const analizGorunur = useMemo(
    () => (ara.trim() ? analiz.filter((a) => eslesir(ara, a.code)) : analiz),
    [analiz, ara]
  );

  return (
    <div>
      <div className="card no-print" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <button className={`btn small ${tab === "fiyat" ? "" : "secondary"}`} onClick={() => setTab("fiyat")}>
          📦 Partiler / Alış Fiyatları
        </button>
        <button className={`btn small ${tab === "analiz" ? "" : "secondary"}`} onClick={() => setTab("analiz")}>
          📈 Satış Analizi
        </button>
        {tab === "analiz" && (
          <input type="month" style={{ width: "auto" }} value={ay} onChange={(e) => setAy(e.target.value)} />
        )}
        <input
          style={{ flex: 1, minWidth: 150 }}
          placeholder="Kod ara…"
          value={ara}
          onChange={(e) => setAra(e.target.value)}
        />
      </div>

      {err && <div className="notice err">{err}</div>}
      {msg && <div className="notice ok">{msg}</div>}

      {tab === "fiyat" && (
        <>
          {/* Parti seçimi */}
          <div className="card" style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
            <div style={{ minWidth: 260, flex: 1 }}>
              <label>Parti (Konteyner)</label>
              <select value={parti?.id || ""} onChange={(e) => setSeciliParti(e.target.value)}>
                {partiler.map((x) => (
                  <option key={x.id} value={x.id}>
                    {x.tarih.split("-").reverse().join(".")} — {x.ad}
                    {x.pct == null ? "  (% girilmedi)" : `  (%${x.pct})`}
                  </option>
                ))}
                {!partiler.length && <option value="">— henüz parti yok —</option>}
              </select>
            </div>
            <button className="btn small" onClick={() => setYpAcik((o) => !o)}>
              {ypAcik ? "Vazgeç" : "+ Yeni Parti"}
            </button>
            {parti && Object.keys(parti.items).length === 0 && (
              <button className="btn small danger" disabled={saving}
                onClick={() => { if (confirm(`"${parti.ad}" partisi silinsin mi?`)) gonder({ partiSil: parti.id }); }}>
                🗑 Partiyi Sil
              </button>
            )}
          </div>

          {ypAcik && (
            <div className="card" style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
              <div style={{ flex: 1, minWidth: 220 }}>
                <label>Parti Adı</label>
                <input value={ypAd} onChange={(e) => setYpAd(e.target.value)}
                  placeholder='örn. "Ağustos 2026 — 40 lık konteyner"' />
              </div>
              <div>
                <label>Geliş Tarihi</label>
                <input type="date" value={ypTarih} onChange={(e) => setYpTarih(e.target.value)} />
              </div>
              <button className="btn" disabled={saving || !ypAd.trim()}
                onClick={async () => {
                  const d = await gonder({ yeniParti: { ad: ypAd.trim(), tarih: ypTarih } });
                  if (d) {
                    const yeni = [...d.partiler].sort((a, b) => b.createdAt.localeCompare(a.createdAt))[0];
                    setSeciliParti(yeni.id);
                    setYpAd(""); setYpAcik(false);
                  }
                }}>
                Parti Aç
              </button>
            </div>
          )}

          {parti && (
            <>
              {/* Partinin yüzdesi */}
              <div className="card" style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
                <div>
                  <label>Bu Partinin Genel Gider %&apos;si</label>
                  <input type="number" step="0.1" min="0" style={{ width: 140 }}
                    value={pctInput} onChange={(e) => setPctInput(e.target.value)}
                    placeholder="sonra girilebilir" />
                </div>
                <button className="btn small" disabled={saving}
                  onClick={() => gonder({ partiId: parti.id, pct: pctInput === "" ? null : Number(pctInput) })}>
                  Yüzdeyi Kaydet
                </button>
                <p style={{ margin: 0, flex: 1, minWidth: 240, fontSize: 12.5,
                  color: parti.pct == null ? "var(--error)" : "var(--muted)" }}>
                  {parti.pct == null
                    ? "Yüzde henüz girilmedi — bu partinin malları için kâr hesaplanmaz. Nakliye/gümrük belli olunca girin."
                    : `Maliyet = alış × ${fmt(1 + parti.pct / 100, 3)}. Nakliye, gümrük, fire dahil.`}
                </p>
              </div>

              {/* Kalem girişi — çerçeve /mt, cam-ayna /m², teknik /kutu */}
              <div className="card" style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
                <div>
                  <label>Tür</label>
                  <select style={{ width: 150 }} value={fTur}
                    onChange={(e) => { setFTur(e.target.value); setFKod(""); }}>
                    {TURLER.map((t) => <option key={t.key} value={t.key}>{t.ad}</option>)}
                  </select>
                </div>
                <div>
                  <label>{fTur === "teknik" ? "Ürün" : fTur === "cam" ? "Cam Türü" : "Ürün Kodu"}</label>
                  {fTur === "cam" ? (
                    <select style={{ width: 170 }} value={fKod} onChange={(e) => setFKod(e.target.value)}>
                      <option value="">Seçin…</option>
                      {CAM_SECENEKLERI.map((n) => <option key={n} value={n}>{n}</option>)}
                    </select>
                  ) : fTur === "teknik" ? (
                    <>
                      <input list="maliyet-teknik" style={{ width: 220 }}
                        value={fKod} onChange={(e) => setFKod(e.target.value)}
                        placeholder="örn. NS Karton Kadife" />
                      <datalist id="maliyet-teknik">
                        {TECHNICAL_PRODUCTS.map((t) => <option key={t.code} value={t.name} />)}
                      </datalist>
                    </>
                  ) : (
                    <>
                      <input list="maliyet-kodlar" style={{ width: 150 }}
                        value={fKod} onChange={(e) => setFKod(e.target.value)} placeholder="örn. 4501 S" />
                      <datalist id="maliyet-kodlar">
                        {FRAME_PROFILES.map((p) => <option key={p.code} value={p.code} />)}
                      </datalist>
                    </>
                  )}
                </div>
                <div>
                  <label>Birim Alış (/{birimEtiket(turBilgi.birim)})</label>
                  <input type="text" inputMode="decimal" style={{ width: 130 }}
                    value={fAlis} onChange={(e) => setFAlis(e.target.value)} />
                </div>
                <div>
                  <label>Para Birimi</label>
                  <select style={{ width: 95 }} value={fBirim} onChange={(e) => setFBirim(e.target.value)}>
                    <option value="USD">$ USD</option>
                    <option value="EUR">€ EUR</option>
                    <option value="TL">₺ TL</option>
                  </select>
                </div>
                <button className="btn"
                  disabled={saving || !fKod.trim() || !(parseFloat(fAlis.replace(",", ".")) > 0)}
                  onClick={() => {
                    gonder({
                      partiId: parti.id,
                      items: [{
                        code: fKod.trim(),
                        alis: parseFloat(fAlis.replace(",", ".")),
                        currency: fBirim,
                        birim: turBilgi.birim,
                      }],
                    });
                    setFKod(""); setFAlis("");
                  }}>
                  + Ekle / Güncelle
                </button>
              </div>

              {/* Kalem listesi */}
              <div className="card" style={{ padding: 0, overflowX: "auto" }}>
                <table>
                  <thead>
                    <tr>
                      <th>Kod / Ürün</th>
                      <th style={{ textAlign: "right" }}>Alış</th>
                      <th style={{ textAlign: "right" }}>Birim Maliyet</th>
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    {kalemler.map((k) => (
                      <tr key={k.code}>
                        <td style={{ fontWeight: 700 }}>{k.code}</td>
                        <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                          {simge(k.currency)}{fmt(k.alis, 4)}
                          <span style={{ color: "var(--muted)", fontWeight: 400 }}>
                            {" "}/{birimEtiket(k.birim)}
                          </span>
                        </td>
                        <td style={{ textAlign: "right", fontWeight: 600 }}>
                          {parti.pct == null
                            ? <span style={{ color: "var(--muted)" }}>% bekliyor</span>
                            : `${simge(k.currency)}${fmt(k.alis * (1 + parti.pct / 100), 4)}`}
                        </td>
                        <td>
                          <button className="btn small danger"
                            onClick={() => { if (confirm(`${k.code} bu partiden silinsin mi?`)) gonder({ partiId: parti.id, sil: [k.code] }); }}>
                            🗑
                          </button>
                        </td>
                      </tr>
                    ))}
                    {!kalemler.length && (
                      <tr><td colSpan={4} style={{ color: "var(--muted)" }}>
                        Bu partiye henüz kalem girilmedi.
                      </td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              <p style={{ fontSize: 12.5, color: "var(--muted)" }}>
                {loading ? "" : `${Object.keys(parti.items).length} kalem · ${parti.ad} · geliş ${parti.tarih.split("-").reverse().join(".")}`}
              </p>
            </>
          )}
        </>
      )}

      {tab === "analiz" && (
        <>
          {ozet && (
            <div className="cari-cards">
              <div className="cari-card">
                <span>Ciro {ay ? `(${ay})` : "(tümü)"}</span>
                <strong>₺{fmt(ozet.toplamCiro)}</strong>
              </div>
              <div className="cari-card">
                <span>Maliyet</span>
                <strong style={{ color: "var(--error)" }}>₺{fmt(ozet.toplamMaliyet)}</strong>
                <span style={{ fontSize: 11.5 }}>maliyeti tam hesaplanan kodlar</span>
              </div>
              <div className="cari-card">
                <span>Kâr</span>
                <strong style={{ color: "var(--success)" }}>₺{fmt(ozet.toplamKar)}</strong>
                {ozet.maliyetliCiro > 0 && (
                  <span style={{ fontSize: 11.5 }}>marj %{fmt((ozet.toplamKar / ozet.maliyetliCiro) * 100, 1)}</span>
                )}
              </div>
              <div className="cari-card">
                <span>Kapsam</span>
                <strong>%{ozet.kapsam}</strong>
                <span style={{ fontSize: 11.5 }}>maliyeti bilinen kod oranı</span>
              </div>
            </div>
          )}
          {loading ? (
            <p style={{ color: "var(--muted)" }}>Hesaplanıyor…</p>
          ) : (
            <div className="card" style={{ padding: 0, overflowX: "auto" }}>
              <table>
                <thead>
                  <tr>
                    <th>Kod / Ürün</th>
                    <th style={{ textAlign: "right" }} title="Çerçevede metre, camda m², teknik malzemede kutu">
                      Satılan (mt/m²/kutu)
                    </th>
                    <th style={{ textAlign: "right" }}>Ciro</th>
                    <th style={{ textAlign: "right" }}>Maliyet</th>
                    <th style={{ textAlign: "right" }}>Kâr</th>
                    <th style={{ textAlign: "right" }}>Marj</th>
                  </tr>
                </thead>
                <tbody>
                  {analizGorunur.map((a) => (
                    <tr key={a.code}>
                      <td style={{ fontWeight: 700 }}>{a.code}</td>
                      <td style={{ textAlign: "right" }}>{fmt(a.metraj)}</td>
                      <td style={{ textAlign: "right" }}>₺{fmt(a.ciro)}</td>
                      <td style={{ textAlign: "right" }}>
                        {a.maliyet != null
                          ? `₺${fmt(a.maliyet)}`
                          : <span style={{ color: "var(--muted)" }}>
                              {a.durum === "yuzde-bekliyor" ? "% bekliyor" : "alış girilmedi"}
                            </span>}
                      </td>
                      <td style={{ textAlign: "right", fontWeight: 600,
                        color: a.kar == null ? undefined : a.kar >= 0 ? "var(--success)" : "var(--error)" }}>
                        {a.kar != null ? `₺${fmt(a.kar)}` : "—"}
                      </td>
                      <td style={{ textAlign: "right", fontWeight: 600,
                        color: a.marj == null ? undefined : a.marj >= 0 ? "var(--success)" : "var(--error)" }}>
                        {a.marj != null ? `%${fmt(a.marj, 1)}` : "—"}
                      </td>
                    </tr>
                  ))}
                  {!analizGorunur.length && (
                    <tr><td colSpan={6} style={{ color: "var(--muted)" }}>
                      Bu dönemde satış satırı yok.
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
          <p style={{ fontSize: 12.5, color: "var(--muted)" }}>
            Her satışın maliyeti, sipariş tarihinden önceki en son partinin
            fiyatından ve o partinin yüzdesinden hesaplanır; kur olarak siparişin
            kendi günlük kuru kullanılır. Renk ekli kodlar taban koda toplanır.
          </p>
        </>
      )}
    </div>
  );
}
