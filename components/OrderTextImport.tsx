"use client";

// WhatsApp/telefon notundan gelen serbest sipariş metnini yapay zekayla
// çözümleyip sipariş formuna satır olarak aktarır.

import { useState } from "react";

export interface ParsedLine {
  kind: "frame" | "glass" | "ayna" | "technical" | "other";
  code: string;
  rawCode: string;
  matched: boolean;
  unit: string;
  qty: number;
  note: string;
  confidence: number;
}

const KIND_LABEL: Record<string, string> = {
  frame: "Çerçeve Profili",
  glass: "Cam",
  ayna: "Ayna",
  technical: "Teknik Malzeme",
  other: "Diğer",
};

const ORNEK = `Merhaba, 3 koli ks2030 beyaz
50 metre gc065-1473
2 kutu 10luk agraf
bir de 122x183 düz cam 4 plaka
Yılmaz Çerçeve, cuma kargoya versin`;

export default function OrderTextImport({
  onApply,
  onClose,
}: {
  onApply: (data: { lines: ParsedLine[]; customer: string; note: string }) => void;
  onClose: () => void;
}) {
  const [text, setText] = useState("");
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");
  const [result, setResult] = useState<{
    lines: ParsedLine[];
    customer: string;
    note: string;
  } | null>(null);
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
      if (!d.lines?.length) {
        setErr("Metinde ürün satırı bulunamadı.");
      } else {
        setResult({ lines: d.lines, customer: d.customer, note: d.note });
        setSelected(new Set(d.lines.map((_: unknown, i: number) => i)));
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
    });
  }

  return (
    <div className="ti-backdrop" onClick={onClose}>
      <div className="ti-modal" onClick={(e) => e.stopPropagation()}>
        <div className="ti-head">
          <b>🤖 Metinden Sipariş Oluştur</b>
          <span style={{ flex: 1 }} />
          <button className="btn small secondary" onClick={onClose}>Kapat</button>
        </div>

        <div className="ti-body">
          <p style={{ fontSize: 13, color: "var(--text-2)", marginBottom: 10 }}>
            Müşteriden gelen WhatsApp mesajını veya telefon notunu olduğu gibi
            yapıştırın; satırlara ayrılıp forma eklenir.
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
              {loading ? "Çözümleniyor..." : "🔍 Çözümle"}
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
              {(result.customer || result.note) && (
                <div className="notice info" style={{ marginTop: 0 }}>
                  {result.customer && <>Müşteri: <strong>{result.customer}</strong></>}
                  {result.customer && result.note && " · "}
                  {result.note && <>Not: {result.note}</>}
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
                    {l.note && <span className="ti-note">{l.note}</span>}
                  </label>
                ))}
              </div>

              <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
                <button className="btn" disabled={selected.size === 0} onClick={apply}>
                  ✓ Seçilenleri Forma Ekle
                </button>
                <button className="btn secondary" onClick={() => setResult(null)}>
                  ← Metni Düzenle
                </button>
              </div>
              <p style={{ fontSize: 11.5, color: "var(--muted)", marginTop: 8 }}>
                Eklenen satırların kod, miktar ve fiyatlarını formda kontrol edin.
              </p>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
