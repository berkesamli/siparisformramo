"use client";

// Sipariş metnini (WhatsApp notu, el yazısından geçirilmiş liste) satırlara
// çevirip forma aktarır. Kod + desen/renk, miktar, birim ve iskonto/KDV
// kural tabanlı okunur (lib/siparis-metin); okunamayan satırlar yapay zekâya
// gider, o da yoksa kullanıcıya gösterilir.

import { useState } from "react";
import { createPortal } from "react-dom";
import Icon from "@/components/shell/Icon";

export interface ParsedLine {
  kind: "frame" | "glass" | "ayna" | "technical" | "other";
  code: string;
  rawCode: string;
  matched: boolean;
  unit: string;
  qty: number;
  note: string;
  confidence: number;
  techCode?: string;
  kartonKodu?: string;
}

export interface ParsedResult {
  lines: ParsedLine[];
  customer: string;
  note: string;
  iskontoPct?: number;
  kdv?: boolean;
}

const KIND_LABEL: Record<string, string> = {
  frame: "Çerçeve Profili",
  glass: "Cam",
  ayna: "Ayna",
  technical: "Teknik Malzeme",
  other: "Diğer",
};

const ORNEK = `Yılmaz Çerçeve / Ankara
%40 isk + KDV
KS 3420-black 10 koli
GB211-4110B 3 koli
3127 S-A79 20 boy
gc065 1473 50 mt
NS 455 → 25 ad
İthal 10'luk agraf 8 kutu
Oluklu karton 50 ad
Araç ile gidecek`;

export default function OrderTextImport({
  onApply,
  onClose,
}: {
  onApply: (data: ParsedResult) => void;
  onClose: () => void;
}) {
  const [text, setText] = useState("");
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");
  const [result, setResult] = useState<(ParsedResult & { okunamayan?: string[] }) | null>(null);
  const [selected, setSelected] = useState<Set<number>>(new Set());

  async function parse() {
    if (!text.trim()) {
      setErr("Önce sipariş metnini yapıştırın.");
      return;
    }
    setLoading(true);
    setErr("");
    setResult(null);
    try {
      const res = await fetch("/api/ai/siparis-coz", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text }),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Çözümlenemedi");
      if (!d.lines?.length && !d.okunamayan?.length) {
        setErr("Metinde ürün satırı bulunamadı.");
      } else {
        setResult({ lines: d.lines || [], customer: d.customer, note: d.note, iskontoPct: d.iskontoPct, kdv: d.kdv, okunamayan: d.okunamayan || [] });
        setSelected(new Set((d.lines || []).map((_: unknown, i: number) => i)));
      }
    } catch (e: any) {
      setErr(e.message || "Bir hata oluştu");
    } finally {
      setLoading(false);
    }
  }

  function toggle(i: number) {
    setSelected((s) => {
      const n = new Set(s);
      if (n.has(i)) n.delete(i);
      else n.add(i);
      return n;
    });
  }

  function apply() {
    if (!result) return;
    onApply({
      lines: result.lines.filter((_, i) => selected.has(i)),
      customer: result.customer,
      note: result.note,
      iskontoPct: result.iskontoPct,
      kdv: result.kdv,
    });
  }

  // Kart (.card) backdrop-filter ile sabit konumlu torunların kapsayıcı bloğu
  // olur — position:fixed karartma kartın içinde kalırdı. Modal bu yüzden
  // portal ile <body>'ye basılır; z-index 120 kabuğun (80–90) üstünde çalışır.
  if (typeof document === "undefined") return null;
  return createPortal(
    <div className="ti-backdrop" onClick={onClose}>
      <div className="ti-modal" onClick={(e) => e.stopPropagation()}>
        <div className="ti-head">
          <span className="card-head-icon">
            <Icon name="sparkles" size={16} />
          </span>
          <b>Metinden Sipariş Oluştur</b>
          <span style={{ flex: 1 }} />
          <button className="btn small secondary" onClick={onClose}>
            <Icon name="x" size={14} />
            Kapat
          </button>
        </div>

        <div className="ti-body">
          <p style={{ fontSize: 13, color: "var(--text-2)", marginBottom: 10 }}>
            Her satıra bir ürün: <code>kod miktar birim</code> (örn. <code>KS 3420-black 10 koli</code>,
            <code>GB211-4110B 20 boy</code>, <code>NS 455 25 ad</code>). Renk/desen eki koda dahil edilir,
            ilk satır müşteri adı, <code>%40 isk + KDV</code> ve teslimat notu da okunur.
            WhatsApp mesajını olduğu gibi yapıştırmak da olur.
          </p>

          <textarea
            rows={7}
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder={ORNEK}
            style={{ resize: "vertical", lineHeight: 1.5 }}
          />

          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <button className="btn" disabled={loading} onClick={parse}>
              {loading ? (
                "Çözümleniyor..."
              ) : (
                <>
                  <Icon name="search" size={16} />
                  Çözümle
                </>
              )}
            </button>
            <button
              className="btn secondary small"
              onClick={() => { setText(ORNEK); setErr(""); }}
            >
              Örnek metni dene
            </button>
          </div>

          {err && <div className="notice err">{err}</div>}

          {result && (
            <>
              <h3 style={{ fontSize: 15, margin: "18px 0 8px" }}>
                Bulunan Satırlar ({selected.size}/{result.lines.length} seçili)
              </h3>
              {(result.customer || result.note || result.iskontoPct !== undefined || result.kdv !== undefined) && (
                <div className="notice info" style={{ marginTop: 0 }}>
                  {[
                    result.customer ? <span key="m">Müşteri: <strong>{result.customer}</strong></span> : null,
                    result.iskontoPct !== undefined ? <span key="i">İskonto: <strong>%{result.iskontoPct}</strong></span> : null,
                    result.kdv !== undefined ? <span key="k">KDV: <strong>{result.kdv ? "var" : "yok"}</strong></span> : null,
                    result.note ? <span key="n">Not: {result.note}</span> : null,
                  ].filter(Boolean).map((el, i, arr) => <span key={i}>{el}{i < arr.length - 1 ? " · " : ""}</span>)}
                </div>
              )}
              {result.okunamayan && result.okunamayan.length > 0 && (
                <div className="notice warn" style={{ marginTop: 0 }}>
                  <strong>Okunamayan satırlar</strong> (elle ekleyin):
                  <ul style={{ margin: "4px 0 0 18px" }}>
                    {result.okunamayan.map((s, i) => <li key={i}><code>{s}</code></li>)}
                  </ul>
                </div>
              )}

              <div className="ti-lines">
                {result.lines.map((l, i) => (
                  <label key={i} className={`ti-line ${selected.has(i) ? "sel" : ""}`}>
                    <input
                      type="checkbox"
                      checked={selected.has(i)}
                      onChange={() => toggle(i)}
                      style={{ width: 18, height: 18, flexShrink: 0 }}
                    />
                    <span className="ti-kind">{KIND_LABEL[l.kind]}</span>
                    <span className="ti-code">
                      {l.code}
                      {l.matched && <i className="ti-ok" title="Katalogda bulundu">✓</i>}
                      {!l.matched && l.kind === "frame" && (
                        <i className="ti-warn" title="Katalogda birebir bulunamadı">?</i>
                      )}
                    </span>
                    <span className="ti-qty">
                      {l.qty} {l.unit}
                    </span>
                    {l.kartonKodu && <span className="ti-note">karton kodu {l.kartonKodu}</span>}
                    {l.note && <span className="ti-note">{l.note}</span>}
                  </label>
                ))}
              </div>

              <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
                <button className="btn" disabled={selected.size === 0} onClick={apply}>
                  <Icon name="check-circle" size={16} />
                  Seçilenleri Forma Ekle
                </button>
                <button className="btn secondary" onClick={() => setResult(null)}>
                  <Icon name="chevron-left" size={16} />
                  Metni Düzenle
                </button>
              </div>
              <p style={{ fontSize: 11.5, color: "var(--muted)", marginTop: 8 }}>
                Eklenen satırların kod, miktar ve fiyatlarını formda kontrol edin.
              </p>
            </>
          )}
        </div>
      </div>
    </div>,
    document.body
  );
}
