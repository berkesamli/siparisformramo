"use client";

import { useRef, useState } from "react";

interface Msg {
  role: "user" | "assistant";
  content: string;
}

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

  if (!open) {
    return (
      <button
        className="btn no-print"
        style={{ position: "fixed", right: 20, bottom: 20, borderRadius: 999, zIndex: 60 }}
        onClick={() => setOpen(true)}
      >
        🤖 Ürün Asistanı
      </button>
    );
  }

  return (
    <div
      className="card no-print"
      style={{
        position: "fixed",
        right: 20,
        bottom: 20,
        width: "min(400px, calc(100vw - 40px))",
        zIndex: 60,
        display: "flex",
        flexDirection: "column",
        maxHeight: "70vh",
        padding: 16,
      }}
    >
      <div style={{ display: "flex", alignItems: "center", marginBottom: 10 }}>
        <strong style={{ flex: 1 }}>🤖 Olga Ürün Asistanı</strong>
        <button className="btn small secondary" onClick={() => setOpen(false)}>
          ✕
        </button>
      </div>
      <div
        ref={listRef}
        style={{ flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: 8, minHeight: 120 }}
      >
        {messages.length === 0 && (
          <p style={{ color: "var(--muted)", fontSize: 13 }}>
            Ürün kodları, liste fiyatları, koli bilgileri veya stok hakkında soru
            sorabilirsiniz. Örn: &quot;KS 2030 fiyatı nedir?&quot;
          </p>
        )}
        {messages.map((m, i) => (
          <div
            key={i}
            style={{
              alignSelf: m.role === "user" ? "flex-end" : "flex-start",
              background: m.role === "user" ? "var(--brand-dark)" : "var(--input)",
              borderRadius: 10,
              padding: "8px 12px",
              maxWidth: "85%",
              fontSize: 13.5,
              whiteSpace: "pre-wrap",
            }}
          >
            {m.content}
          </div>
        ))}
        {loading && (
          <div style={{ color: "var(--muted)", fontSize: 13 }}>Yazıyor…</div>
        )}
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && send()}
          placeholder="Sorunuzu yazın…"
        />
        <button className="btn small" onClick={send} disabled={loading}>
          Gönder
        </button>
      </div>
    </div>
  );
}
