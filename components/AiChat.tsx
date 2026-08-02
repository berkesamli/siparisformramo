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
          <div style={{ color: "var(--muted)", fontSize: 13 }}>
            <p style={{ margin: "0 0 8px" }}>
              Fiyat, güncel stok, kur, perakende çerçeveletme hesabı, sipariş ve
              müşteri kayıtları hakkında soru sorabilirsiniz.
            </p>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {[
                "GC065 stokta kaç boy var?",
                "Bugünün kuru ne?",
                "50x70 eser, GB139, 5 cm paspartu, mat cam kaça olur?",
                "Yılmaz Çerçeve'nin açık bakiyesi ne kadar?",
              ].map((s) => (
                <button
                  key={s}
                  className="btn small secondary"
                  style={{ fontSize: 11.5, padding: "4px 9px" }}
                  onClick={() => setInput(s)}
                >
                  {s}
                </button>
              ))}
            </div>
          </div>
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
