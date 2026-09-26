"use client";

import { useEffect, useRef, useState } from "react";
import { konusmaTanimaDestegi, konusmaTanimaOlustur, seslendirmeMetni, tarayiciSesi, type SpeechRecognitionBenzeri } from "@/lib/sesli-asistan";
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

  // ---- Jarvis: sesli giriş / sesli yanıt ----
  // Mikrofon: tarayıcının konuşma tanıması (tr-TR). Sesli yanıt: sunucu sesi (ElevenLabs tanımlıysa),
  // yoksa tarayıcının Türkçe sesi. "Sürekli konuşma" açıkken yanıt bitince mikrofon kendiliğinden açılır.
  const [sesDestegi, setSesDestegi] = useState(false);
  const [dinliyor, setDinliyor] = useState(false);
  const [konusuyor, setKonusuyor] = useState(false);
  const [sesliYanit, setSesliYanit] = useState(false);
  const [surekli, setSurekli] = useState(false);
  const [araMetin, setAraMetin] = useState("");
  const [sesHata, setSesHata] = useState("");
  const tanimaRef = useRef<SpeechRecognitionBenzeri | null>(null);
  const sesRef = useRef<HTMLAudioElement | null>(null);
  const sunucuSesi = useRef<boolean | null>(null); // null: denenmedi, false: yok (501), true: var
  const surekliRef = useRef(false);
  const loadingRef = useRef(false);
  surekliRef.current = surekli;
  loadingRef.current = loading;

  useEffect(() => { setSesDestegi(konusmaTanimaDestegi()); }, []);
  useEffect(() => () => durdur(), []); // eslint-disable-line react-hooks/exhaustive-deps

  function durdur() {
    try { tanimaRef.current?.abort(); } catch { /* yok */ }
    tanimaRef.current = null;
    setDinliyor(false);
    setAraMetin("");
    if (sesRef.current) { try { sesRef.current.pause(); } catch { /* yok */ } sesRef.current = null; }
    if (typeof window !== "undefined" && "speechSynthesis" in window) window.speechSynthesis.cancel();
    setKonusuyor(false);
  }

  function dinle() {
    const rec = konusmaTanimaOlustur();
    if (!rec) { setSesHata("Bu tarayıcı sesli girişi desteklemiyor; Chrome, Edge ya da Safari kullanın."); return; }
    durdur();
    setSesHata("");
    rec.lang = "tr-TR";
    rec.interimResults = true;
    rec.continuous = false;
    rec.maxAlternatives = 1;
    let son = "";
    rec.onresult = (e) => {
      let kesin = "", gecici = "";
      for (let i = 0; i < e.results.length; i++) { const r = e.results[i]; if (r.isFinal) kesin += r[0].transcript; else gecici += r[0].transcript; }
      setAraMetin(gecici);
      if (kesin.trim()) son = kesin.trim();
    };
    rec.onerror = (e) => {
      if (e.error === "not-allowed" || e.error === "service-not-allowed") { setSesHata("Mikrofon izni verilmedi. Tarayıcının adres çubuğundan mikrofona izin verin."); setSurekli(false); }
      else if (e.error === "network") setSesHata("Konuşma tanıma servisine ulaşılamadı (ağ).");
      // "no-speech" / "aborted": sessizce geç
    };
    rec.onend = () => {
      tanimaRef.current = null;
      setDinliyor(false);
      setAraMetin("");
      if (son) { setSesliYanit(true); void send(son, true); }
      else if (surekliRef.current && !loadingRef.current) setTimeout(() => { if (surekliRef.current && !tanimaRef.current && !loadingRef.current) dinle(); }, 400);
    };
    tanimaRef.current = rec;
    setDinliyor(true);
    try { rec.start(); } catch { setDinliyor(false); tanimaRef.current = null; }
  }

  async function seslendir(metin: string) {
    const okunacak = seslendirmeMetni(metin);
    if (!okunacak) return;
    setKonusuyor(true);
    try {
      if (sunucuSesi.current !== false) {
        const r = await fetch("/api/ai/ses", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ metin: okunacak }) });
        if (r.ok && (r.headers.get("content-type") || "").includes("audio")) {
          sunucuSesi.current = true;
          const url = URL.createObjectURL(await r.blob());
          const a = new Audio(url);
          sesRef.current = a;
          await new Promise<void>((cozul) => { a.onended = () => cozul(); a.onerror = () => cozul(); a.play().catch(() => cozul()); });
          URL.revokeObjectURL(url);
          return;
        }
        if (r.status === 501) sunucuSesi.current = false; // sunucu sesi tanımlı değil: bundan sonra tarayıcı sesi
      }
      await tarayiciSesi(okunacak);
    } catch { /* ses çalınamadı: metin zaten ekranda */ }
    finally {
      sesRef.current = null;
      setKonusuyor(false);
    }
  }

  async function send(metin?: string, sesli = false) {
    const text = (metin ?? input).trim();
    if (!text || loadingRef.current) return;
    const sesliMod = sesli || sesliYanit;
    const next: Msg[] = [...messages, { role: "user", content: text }];
    setMessages(next);
    setInput("");
    setLoading(true);
    loadingRef.current = true;
    let cevap = "";
    try {
      const res = await fetch("/api/ai", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: next, sesli: sesliMod }),
      });
      const data = await res.json();
      cevap = data.ok ? data.reply : `⚠️ ${data.error || "Hata oluştu."}`;
      setMessages([...next, { role: "assistant", content: cevap }]);
    } catch {
      cevap = "⚠️ Sunucuya ulaşılamadı.";
      setMessages([...next, { role: "assistant", content: cevap }]);
    } finally {
      setLoading(false);
      loadingRef.current = false;
      setTimeout(() => {
        listRef.current?.scrollTo({ top: 99999, behavior: "smooth" });
      }, 50);
    }
    if (sesliMod && cevap) await seslendir(cevap);
    // Sürekli konuşma: yanıt bittikten sonra yeniden dinle
    if (surekliRef.current && !tanimaRef.current) dinle();
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
        <Icon name="sparkles" size={16} /> Jarvis
      </button>
    );
  }

  return (
    <div
      className="card pad-sm no-print"
      role="dialog"
      aria-label="Jarvis — Olga Ürün Asistanı"
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
          <h2>Jarvis</h2>
          <div className="card-head-sub">Olga ürün ve sipariş asistanı</div>
        </div>
        <span className="spacer" />
        <div className="card-head-actions">
          <button
            type="button"
            className={`btn icon small ${sesliYanit ? "" : "ghost"}`}
            aria-pressed={sesliYanit}
            aria-label={sesliYanit ? "Sesli yanıtı kapat" : "Sesli yanıtı aç"}
            title={sesliYanit ? "Sesli yanıt açık — kapat" : "Yanıtları sesli oku"}
            onClick={() => { if (sesliYanit) { durdur(); setSurekli(false); } setSesliYanit(!sesliYanit); }}
          >
            <Icon name={sesliYanit ? "volume" : "volume-off"} size={16} />
          </button>
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

      {(dinliyor || konusuyor || sesHata || surekli) && (
        <div className="ai-durum" aria-live="polite">
          {dinliyor ? (
            <><span className="ai-nokta dinliyor" /> Dinliyorum… {araMetin && <em>{araMetin}</em>}</>
          ) : konusuyor ? (
            <><span className="ai-nokta konusuyor" /> Jarvis konuşuyor… <button type="button" className="btn ghost xs" onClick={durdur}>Durdur</button></>
          ) : sesHata ? (
            <span style={{ color: "var(--warning)" }}>{sesHata}</span>
          ) : (
            <><span className="ai-nokta" /> Sürekli konuşma açık: yanıttan sonra mikrofon yeniden açılır.</>
          )}
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 8, flex: "0 0 auto", alignItems: "center" }}>
        <button
          type="button"
          className={`btn icon small ai-mic ${dinliyor ? "dinliyor" : "secondary"}`}
          aria-pressed={dinliyor}
          aria-label={dinliyor ? "Dinlemeyi durdur" : "Sesle sor"}
          title={sesDestegi ? (dinliyor ? "Dinlemeyi durdur" : "Sesle sor (Türkçe)") : "Bu tarayıcı sesli girişi desteklemiyor; Chrome, Edge ya da Safari kullanın"}
          disabled={!sesDestegi || loading}
          onClick={() => (dinliyor ? durdur() : dinle())}
        >
          <Icon name="mic" size={16} />
        </button>
        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && send()}
          placeholder={dinliyor ? "Konuşun…" : "Sorunuzu yazın ya da mikrofona basın…"}
          aria-label="Sorunuzu yazın"
        />
        <button type="button" className="btn small" onClick={() => send()} disabled={loading}>
          Gönder
        </button>
      </div>
      {sesDestegi && (
        <label className="ai-surekli">
          <input type="checkbox" checked={surekli} onChange={(e) => { const v = e.target.checked; setSurekli(v); if (v) { setSesliYanit(true); if (!dinliyor && !loading && !konusuyor) dinle(); } else durdur(); }} />
          Sürekli konuşma (Jarvis her yanıttan sonra dinler)
        </label>
      )}
    </div>
  );
}
