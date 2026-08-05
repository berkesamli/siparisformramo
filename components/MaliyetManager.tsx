"use client";

// Alış fiyatları + kod bazlı satış/kâr analizi — yalnızca firma sahipleri.
// Maliyet = alış × (1 + genel gider %). Analizde kur olarak her siparişin
// kendi günlük kuru kullanılır.

import { useCallback, useEffect, useMemo, useState } from "react";
import { FRAME_PROFILES } from "@/data/catalog";
import type { MaliyetData } from "@/lib/maliyet";
import { eslesir } from "@/lib/search-norm";

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
  satirSayisi: number;
}

interface Ozet {
  toplamCiro: number;
  maliyetliCiro: number;
  toplamMaliyet: number;
  toplamKar: number;
  kapsam: number;
}

export default function MaliyetManager() {
  const [tab, setTab] = useState<"fiyat" | "analiz">("fiyat");
  const [data, setData] = useState<MaliyetData | null>(null);
  const [analiz, setAnaliz] = useState<AnalizSatir[]>([]);
  const [ozet, setOzet] = useState<Ozet | null>(null);
  const [ay, setAy] = useState("");
  const [ara, setAra] = useState("");
  const [err, setErr] = useState("");
  const [msg, setMsg] = useState("");
  const [loading, setLoading] = useState(true);

  // giriş formu
  const [fKod, setFKod] = useState("");
  const [fAlis, setFAlis] = useState("");
  const [fBirim, setFBirim] = useState("USD");
  const [fPct, setFPct] = useState("");
  const [genelPct, setGenelPct] = useState("");
  const [saving, setSaving] = useState(false);

  const load = useCallback(() => {
    setLoading(true);
    const url =
      tab === "analiz"
        ? `/api/maliyet?analiz=1${ay ? `&ay=${ay}` : ""}`
        : "/api/maliyet";
    fetch(url)
      .then((r) => r.json())
      .then((d) => {
        if (!d.ok) {
          setErr(d.error || "Yüklenemedi");
          return;
        }
        setData(d.data);
        setGenelPct(String(d.data.defaultPct ?? 0));
        if (d.analiz) {
          setAnaliz(d.analiz);
          setOzet(d.ozet);
        }
        setErr("");
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [tab, ay]);

  useEffect(() => {
    load();
  }, [load]);

  async function kaydet(items?: object[], sil?: string[], defaultPct?: number) {
    setSaving(true);
    setMsg("");
    try {
      const r = await fetch("/api/maliyet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items, sil, defaultPct }),
      });
      const d = await r.json();
      if (!d.ok) throw new Error(d.error);
      setData(d.data);
      setMsg("Kaydedildi.");
      setTimeout(() => setMsg(""), 2500);
    } catch (e) {
      setErr(e instanceof Error ? e.message : "Kaydedilemedi");
    } finally {
      setSaving(false);
    }
  }

  const kayitlar = useMemo(() => {
    const list = Object.values(data?.items || {});
    list.sort((a, b) => a.code.localeCompare(b.code, "tr"));
    return ara.trim() ? list.filter((k) => eslesir(ara, k.code, k.note)) : list;
  }, [data, ara]);

  const analizGorunur = useMemo(
    () => (ara.trim() ? analiz.filter((a) => eslesir(ara, a.code)) : analiz),
    [analiz, ara]
  );

  const pctGoster = (pct?: number) =>
    pct != null ? pct : Number(genelPct) || data?.defaultPct || 0;

  return (
    <div>
      <div className="card no-print" style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <button className={`btn small ${tab === "fiyat" ? "" : "secondary"}`} onClick={() => setTab("fiyat")}>
          Alış Fiyatları ({Object.keys(data?.items || {}).length})
        </button>
        <button className={`btn small ${tab === "analiz" ? "" : "secondary"}`} onClick={() => setTab("analiz")}>
          📈 Satış Analizi
        </button>
        {tab === "analiz" && (
          <input type="month" style={{ width: "auto" }} value={ay} onChange={(e) => setAy(e.target.value)} />
        )}
        <input
          style={{ flex: 1, minWidth: 160 }}
          placeholder="Kod ara…"
          value={ara}
          onChange={(e) => setAra(e.target.value)}
        />
      </div>

      {err && <div className="notice err">{err}</div>}
      {msg && <div className="notice ok">{msg}</div>}

      {tab === "fiyat" && (
        <>
          {/* Genel gider yüzdesi */}
          <div className="card" style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
            <div>
              <label>Genel Gider Yüzdesi (%)</label>
              <input
                type="number" step="0.1" min="0" style={{ width: 140 }}
                value={genelPct}
                onChange={(e) => setGenelPct(e.target.value)}
              />
            </div>
            <button
              className="btn small"
              disabled={saving}
              onClick={() => kaydet(undefined, undefined, Number(genelPct) || 0)}
            >
              Yüzdeyi Kaydet
            </button>
            <p style={{ margin: 0, flex: 1, minWidth: 240, fontSize: 12.5, color: "var(--muted)" }}>
              Maliyet = alış × (1 + %{fmt(Number(genelPct) || 0, 1)}). Nakliye, gümrük,
              fire ve işçilik payını kapsar; koda özel yüzde girilirse o geçerli olur.
            </p>
          </div>

          {/* Yeni kayıt */}
          <div className="card" style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
            <div>
              <label>Ürün Kodu</label>
              <input
                list="maliyet-kodlar" style={{ width: 150 }}
                value={fKod} onChange={(e) => setFKod(e.target.value)}
                placeholder="örn. 4501 S"
              />
              <datalist id="maliyet-kodlar">
                {FRAME_PROFILES.map((p) => (
                  <option key={p.code} value={p.code} />
                ))}
              </datalist>
            </div>
            <div>
              <label>Alış Fiyatı</label>
              <input
                type="number" step="0.0001" min="0" style={{ width: 120 }}
                value={fAlis} onChange={(e) => setFAlis(e.target.value)}
                placeholder="/mt"
              />
            </div>
            <div>
              <label>Birim</label>
              <select style={{ width: 90 }} value={fBirim} onChange={(e) => setFBirim(e.target.value)}>
                <option value="USD">$ USD</option>
                <option value="EUR">€ EUR</option>
                <option value="TL">₺ TL</option>
              </select>
            </div>
            <div>
              <label>Özel % (boş = genel)</label>
              <input
                type="number" step="0.1" min="0" style={{ width: 120 }}
                value={fPct} onChange={(e) => setFPct(e.target.value)}
              />
            </div>
            <button
              className="btn"
              disabled={saving || !fKod.trim() || !(parseFloat(fAlis) > 0)}
              onClick={() => {
                kaydet([{ code: fKod.trim(), alis: parseFloat(fAlis), currency: fBirim,
                          pct: fPct === "" ? undefined : parseFloat(fPct) }]);
                setFKod(""); setFAlis(""); setFPct("");
              }}
            >
              + Ekle / Güncelle
            </button>
          </div>

          {loading ? (
            <p style={{ color: "var(--muted)" }}>Yükleniyor…</p>
          ) : (
            <div className="card" style={{ padding: 0, overflowX: "auto" }}>
              <table>
                <thead>
                  <tr>
                    <th>Kod</th>
                    <th style={{ textAlign: "right" }}>Alış</th>
                    <th style={{ textAlign: "right" }}>%</th>
                    <th style={{ textAlign: "right" }}>Birim Maliyet</th>
                    <th>Güncelleyen</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {kayitlar.map((k) => {
                    const pct = pctGoster(k.pct);
                    const maliyet = k.alis * (1 + pct / 100);
                    return (
                      <tr key={k.code}>
                        <td style={{ fontWeight: 700 }}>{k.code}</td>
                        <td style={{ textAlign: "right" }}>
                          {simge(k.currency)}{fmt(k.alis, 4)}
                        </td>
                        <td style={{ textAlign: "right", color: k.pct != null ? "var(--brand)" : "var(--muted)" }}>
                          %{fmt(pct, 1)}{k.pct != null ? " (özel)" : ""}
                        </td>
                        <td style={{ textAlign: "right", fontWeight: 600 }}>
                          {simge(k.currency)}{fmt(maliyet, 4)}
                        </td>
                        <td style={{ fontSize: 12 }}>{k.by}</td>
                        <td>
                          <button className="btn small danger" onClick={() => {
                            if (confirm(`${k.code} alış kaydı silinsin mi?`)) kaydet(undefined, [k.code]);
                          }}>🗑</button>
                        </td>
                      </tr>
                    );
                  })}
                  {!kayitlar.length && (
                    <tr>
                      <td colSpan={6} style={{ color: "var(--muted)" }}>
                        Henüz alış fiyatı girilmedi. Yukarıdan kod seçip ekleyin.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
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
                <span style={{ fontSize: 11.5 }}>alışı girilen kodlar</span>
              </div>
              <div className="cari-card">
                <span>Kâr</span>
                <strong style={{ color: "var(--success)" }}>₺{fmt(ozet.toplamKar)}</strong>
                {ozet.maliyetliCiro > 0 && (
                  <span style={{ fontSize: 11.5 }}>
                    marj %{fmt((ozet.toplamKar / ozet.maliyetliCiro) * 100, 1)}
                  </span>
                )}
              </div>
              <div className="cari-card">
                <span>Maliyet Kapsamı</span>
                <strong>%{ozet.kapsam}</strong>
                <span style={{ fontSize: 11.5 }}>kodların alışı girilmiş oranı</span>
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
                    <th>Kod</th>
                    <th style={{ textAlign: "right" }}>Satılan (mt)</th>
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
                        {a.maliyet != null ? `₺${fmt(a.maliyet)}` : <span style={{ color: "var(--muted)" }}>alış girilmedi</span>}
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
                    <tr>
                      <td colSpan={6} style={{ color: "var(--muted)" }}>
                        Bu dönemde satış satırı yok.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
          <p style={{ fontSize: 12.5, color: "var(--muted)" }}>
            Maliyet, her siparişin kendi günlük kuru ile TL&apos;ye çevrilir. Renk ekli
            kodlar (4501S-1242) taban koda (4501 S) toplanır.
          </p>
        </>
      )}
    </div>
  );
}
