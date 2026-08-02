"use client";

// Bayi sipariş talebi — müşteri portalında kullanılır.
// Bayi ürün kodu + miktar girer, talep çalışana "onay bekliyor" olarak düşer.
// Fiyatlandırma çalışan tarafında yapıldığı için burada tutar gösterilmez.

import { useEffect, useState } from "react";
import { findProfile } from "@/data/catalog";
import { REQUEST_LABELS, type SavedRequest } from "@/lib/requests";

interface Line {
  id: number;
  code: string;
  unit: string;
  qty: string;
  note: string;
}

let seq = 1;
const emptyLine = (): Line => ({ id: seq++, code: "", unit: "Metre", qty: "", note: "" });

export default function DealerRequestForm({ userName }: { userName: string }) {
  const [lines, setLines] = useState<Line[]>([emptyLine()]);
  const [customer, setCustomer] = useState(userName);
  const [phone, setPhone] = useState("");
  const [note, setNote] = useState("");
  const [sending, setSending] = useState(false);
  const [msg, setMsg] = useState("");
  const [err, setErr] = useState("");
  const [history, setHistory] = useState<SavedRequest[]>([]);

  async function loadHistory() {
    try {
      const res = await fetch("/api/talepler");
      if (res.ok) {
        const d = await res.json();
        setHistory(d.requests || []);
      }
    } catch {
      /* geçmiş getirilemezse form yine çalışsın */
    }
  }

  useEffect(() => {
    loadHistory();
  }, []);

  function update(id: number, patch: Partial<Line>) {
    setLines((ls) => ls.map((l) => (l.id === id ? { ...l, ...patch } : l)));
  }

  async function submit() {
    const valid = lines.filter((l) => l.code.trim() && Number(l.qty) > 0);
    if (valid.length === 0) {
      setErr("En az bir ürün kodu ve miktarı girin.");
      return;
    }
    setSending(true);
    setErr("");
    setMsg("");
    try {
      const res = await fetch("/api/talepler", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          customer: customer.trim(),
          phone: phone.trim(),
          note: note.trim(),
          lines: valid.map((l) => ({
            code: l.code.trim().toUpperCase(),
            unit: l.unit,
            qty: Number(l.qty),
            note: l.note.trim(),
          })),
        }),
      });
      const d = await res.json();
      if (!res.ok || !d.ok) throw new Error(d.error || "Talep gönderilemedi");
      setMsg(
        "Talebiniz iletildi. Fiyat ve teyit için en kısa sürede size dönüş yapılacaktır."
      );
      setLines([emptyLine()]);
      setNote("");
      loadHistory();
    } catch (e: any) {
      setErr(e.message || "Bir hata oluştu");
    } finally {
      setSending(false);
    }
  }

  return (
    <div>
      <div className="card">
        <h2 style={{ marginTop: 0 }}>🛒 Sipariş Talebi Oluştur</h2>
        <p style={{ color: "var(--text-2)", fontSize: 13.5, marginTop: -6, marginBottom: 16 }}>
          Ürün kodu ve miktarı girin; talebiniz ekibimize düşer, fiyat teyidiyle
          birlikte size dönüş yapılır.
        </p>

        <div className="rw-grid2">
          <div>
            <label>Firma / Ad</label>
            <input value={customer} onChange={(e) => setCustomer(e.target.value)} />
          </div>
          <div>
            <label>Telefon</label>
            <input value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="05xx xxx xx xx" />
          </div>
        </div>

        <h3 style={{ fontSize: 15, margin: "20px 0 10px" }}>Ürünler</h3>
        {lines.map((l, i) => {
          const p = l.code.trim() ? findProfile(l.code.trim()) : undefined;
          return (
            <div key={l.id} className="req-line">
              <div style={{ flex: "1 1 170px" }}>
                <label>Ürün / Profil Kodu</label>
                <input
                  value={l.code}
                  onChange={(e) => update(l.id, { code: e.target.value.toUpperCase() })}
                  placeholder="örn. KS 2030"
                />
                {l.code.trim() && (
                  <span style={{ fontSize: 11.5, color: p ? "var(--success)" : "var(--muted)" }}>
                    {p ? `✓ ${p.code} · ${p.series} serisi` : "Katalogda bulunamadı — yine de gönderebilirsiniz"}
                  </span>
                )}
              </div>
              <div style={{ width: 120 }}>
                <label>Birim</label>
                <select value={l.unit} onChange={(e) => update(l.id, { unit: e.target.value })}>
                  <option>Metre</option>
                  <option>Koli</option>
                  <option>Adet</option>
                </select>
              </div>
              <div style={{ width: 110 }}>
                <label>Miktar</label>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={l.qty}
                  onChange={(e) => update(l.id, { qty: e.target.value })}
                />
              </div>
              <div style={{ flex: "1 1 150px" }}>
                <label>Not</label>
                <input
                  value={l.note}
                  onChange={(e) => update(l.id, { note: e.target.value })}
                  placeholder="renk, açıklama…"
                />
              </div>
              <button
                className="btn small danger"
                style={{ alignSelf: "flex-end", marginBottom: 1 }}
                disabled={lines.length === 1}
                onClick={() => setLines((ls) => ls.filter((x) => x.id !== l.id))}
              >
                Sil
              </button>
            </div>
          );
        })}

        <button
          className="btn secondary small"
          style={{ marginTop: 6 }}
          onClick={() => setLines((ls) => [...ls, emptyLine()])}
        >
          + Satır Ekle
        </button>

        <div style={{ marginTop: 16 }}>
          <label>Sipariş Notu</label>
          <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Teslimat / açıklama (opsiyonel)" />
        </div>

        <div style={{ marginTop: 18 }}>
          <button className="btn" disabled={sending} onClick={submit}>
            {sending ? "Gönderiliyor..." : "📨 Talebi Gönder"}
          </button>
        </div>

        {msg && <div className="notice ok">{msg}</div>}
        {err && <div className="notice err">{err}</div>}
      </div>

      {history.length > 0 && (
        <>
          <h2>Önceki Taleplerim</h2>
          <div className="card" style={{ padding: 0, overflowX: "auto" }}>
            <table>
              <thead>
                <tr>
                  <th>Tarih</th>
                  <th>Ürünler</th>
                  <th>Durum</th>
                </tr>
              </thead>
              <tbody>
                {history.map((r) => (
                  <tr key={r.id}>
                    <td style={{ whiteSpace: "nowrap" }}>
                      {new Date(r.createdAt).toLocaleDateString("tr-TR")}
                    </td>
                    <td style={{ fontSize: 13 }}>
                      {r.lines.map((l) => `${l.code} ${l.qty} ${l.unit}`).join(", ")}
                    </td>
                    <td>
                      <span
                        className={`badge ${
                          r.status === "onaylandi" ? "var" : r.status === "reddedildi" ? "yok" : "az"
                        }`}
                      >
                        {REQUEST_LABELS[r.status]}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}
