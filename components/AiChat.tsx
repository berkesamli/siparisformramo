"use client";

import { useRef, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Msg {
  role: "user" | "assistant";
  content: string;
}

const ONERILER = [
  "GC065 stokta kaç boy var?",
  "Bugünün kuru ne?",
  "50x70 eser, GB139, 5 cm paspartu, mat cam kaça olur?",
  "Yılmaz Çerçeve'nin Mikro bakiyesi ne kadar?",
];

export default function AiChat() {
  const [open, setOpen] = useState(false);
  const [messages, setMessages] = useState<Msg[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const listRef = useRef<HTMLDivElement>(null);

  async function send() {
    const text = input.trim();
    if (!text || loading) return;
    const next: Msg[] = [...messages, { role: "user", content: text }];
    setMessages(next);
    setInput("");
    setLoading(true);
    try {
      const res = await fetch("/api/ai", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: next }),
      });
      const data = await res.json();
      setMessages([
        ...next,
        {
          role: "assistant",
          content: data.ok ? data.reply : `⚠️ ${data.error || "Hata oluştu."}`,
        },
      ]);
    } catch {
      setMessages([
        ...next,
        { role: "assistant", content: "⚠️ Sunucuya ulaşılamadı." },
      ]);
    } finally {
      setLoading(false);
      setTimeout(() => {
        listRef.current?.scrollTo({ top: 99999, behavior: "smooth" });
      }, 50);
    }
  }

  // Yüzen düğme: sağ altta, telefonda alt sekme çubuğunun üstünde kalır
  // (--fab-bottom kabuk tarafından ayarlanır). Sol alt köşe boş bırakılır.
  if (!open) {
    return (
      <button
        type="button"
        className="btn no-print"
        style={{
          position: "fixed",
          right: 20,
          bottom: "var(--fab-bottom, 20px)",
          borderRadius: 999,
          zIndex: 60,
          boxShadow: "var(--shadow-md)",
        }}
        onClick={() => setOpen(true)}
      >
        <Icon name="sparkles" size={16} /> Ürün Asistanı
      </button>
    );
  }

  return (
    <div
      className="card pad-sm no-print"
      role="dialog"
      aria-label="Olga Ürün Asistanı"
      style={{
        position: "fixed",
        right: 20,
        bottom: "var(--fab-bottom, 20px)",
        width: "min(400px, calc(100vw - 40px))",
        maxHeight: "min(70vh, 640px)",
        zIndex: 60,
        display: "flex",
        flexDirection: "column",
        boxShadow: "var(--shadow-lg)",
      }}
    >
      <div className="card-head" style={{ flex: "0 0 auto" }}>
        <span className="card-head-icon">
          <Icon name="sparkles" size={16} />
        </span>
        <div>
          <h2>Olga Ürün Asistanı</h2>
        </div>
        <span className="spacer" />
        <div className="card-head-actions">
          <button
            type="button"
            className="btn icon small ghost"
            aria-label="Kapat"
            title="Kapat"
            onClick={() => setOpen(false)}
          >
            <Icon name="x" size={16} />
          </button>
        </div>
      </div>

      <div
        ref={listRef}
        style={{
          flex: "1 1 auto",
          overflowY: "auto",
          display: "flex",
          flexDirection: "column",
          gap: 8,
          minHeight: 120,
        }}
      >
        {messages.length === 0 && (
          <div style={{ color: "var(--muted)", fontSize: 13 }}>
            <p style={{ margin: "0 0 8px" }}>
              Fiyat, güncel stok, kur, perakende çerçeveletme hesabı, sipariş ve
              müşteri kayıtları hakkında soru sorabilirsiniz.
            </p>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {ONERILER.map((s) => (
                <button
                  key={s}
                  type="button"
                  className="btn xs secondary"
                  style={{ whiteSpace: "normal", textAlign: "left" }}
                  onClick={() => setInput(s)}
                >
                  {s}
                </button>
              ))}
            </div>
          </div>
        )}
        {messages.map((m, i) => {
          const user = m.role === "user";
          return (
            <div
              key={i}
              style={{
                alignSelf: user ? "flex-end" : "flex-start",
                background: user ? "var(--brand-dark)" : "var(--surface-2)",
                color: user ? "var(--on-brand)" : "var(--text)",
                border: `1px solid ${user ? "transparent" : "var(--border)"}`,
                borderRadius: "var(--radius-sm)",
                padding: "8px 12px",
                maxWidth: "85%",
                fontSize: 13.5,
                whiteSpace: "pre-wrap",
                overflowWrap: "anywhere",
              }}
            >
              {m.content}
            </div>
          );
        })}
        {loading && (
          <div style={{ color: "var(--muted)", fontSize: 13 }}>Yazıyor…</div>
        )}
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 10, flex: "0 0 auto" }}>
        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && send()}
          placeholder="Sorunuzu yazın…"
          aria-label="Sorunuzu yazın"
        />
        <button type="button" className="btn small" onClick={send} disabled={loading}>
          Gönder
        </button>
      </div>
    </div>
  );
}
