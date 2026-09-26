// Jarvis — paneldeki sesli asistan yardımcıları (istemci tarafı; saf fonksiyonlar test edilir).
//
// Konuşma → metin: tarayıcının Web Speech API'si (Chrome, Edge, Safari; Firefox yok), dil tr-TR.
// Metin → ses: önce sunucu (/api/ai/ses, ElevenLabs tanımlıysa doğal ses), yoksa tarayıcının
// kendi Türkçe sesi (speechSynthesis). Sesli okuma için metin markdown/emoji'den arındırılır,
// kod ve para birimleri okunur biçime çevrilir.

export function konusmaTanimaDestegi(): boolean {
  if (typeof window === "undefined") return false;
  const w = window as unknown as { SpeechRecognition?: unknown; webkitSpeechRecognition?: unknown };
  return Boolean(w.SpeechRecognition || w.webkitSpeechRecognition);
}

export function konusmaTanimaOlustur(): SpeechRecognitionBenzeri | null {
  if (typeof window === "undefined") return null;
  const w = window as unknown as { SpeechRecognition?: new () => SpeechRecognitionBenzeri; webkitSpeechRecognition?: new () => SpeechRecognitionBenzeri };
  const Sinif = w.SpeechRecognition || w.webkitSpeechRecognition;
  return Sinif ? new Sinif() : null;
}

/** Web Speech API'nin kullandığımız kadarı (TS lib'de tanımlı değil). */
export interface SpeechRecognitionBenzeri {
  lang: string;
  interimResults: boolean;
  continuous: boolean;
  maxAlternatives: number;
  onresult: ((e: { results: ArrayLike<{ isFinal: boolean; 0: { transcript: string } }> }) => void) | null;
  onerror: ((e: { error: string }) => void) | null;
  onend: (() => void) | null;
  start(): void;
  stop(): void;
  abort(): void;
}

/**
 * Sesli okuma metni: markdown/emoji/bağlantı kalıntıları atılır; para birimi, yüzde ve
 * ürün kodları okunur hâle gelir ("₺3.913,44" → "3.913,44 lira", "KS4022-BİG" → "KS 4022 BİG").
 * En çok 900 karakter (ses servisleri karakter başına ücretlendirir).
 */
export function seslendirmeMetni(metin: string, max = 900): string {
  let s = String(metin || "")
    .replace(/```[\s\S]*?```/g, " ")
    .replace(/`([^`]*)`/g, "$1")
    .replace(/!\[[^\]]*\]\([^)]*\)/g, " ")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
    .replace(/https?:\/\/\S+/g, " ")
    .replace(/^\s{0,3}#{1,6}\s+/gm, "")
    .replace(/^\s*[-*•]\s+/gm, "")
    .replace(/^\s*\d+[.)]\s+/gm, "")
    .replace(/\*\*|__|~~|\*|_/g, "")
    .replace(/\|/g, ", ")
    .replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}\u{200D}]/gu, "")
    .replace(/[⚠️🤖✅❌➡️→←↔]/g, "")
    .replace(/≈/g, "yaklaşık ")
    .replace(/·/g, ", ")
    .replace(/₺\s?([\d.,]+)/g, "$1 lira")
    .replace(/\$\s?([\d.,]+)/g, "$1 dolar")
    .replace(/€\s?([\d.,]+)/g, "$1 euro")
    .replace(/%\s?([\d.,]+)/g, "yüzde $1")
    .replace(/\/\s?mt\b/g, " metre başına")
    .replace(/\bmt\b/g, "metre")
    .replace(/\bm²/g, "metrekare");
  // Ürün kodları: harf-rakam sınırlarına boşluk, tireyi boşluk yap (KS4022-BİG → KS 4022 BİG, 4501S-1242 → 4501 S 1242)
  s = s.replace(/\b([A-ZÇĞİÖŞÜ]{1,4})(\d{2,})/g, "$1 $2").replace(/(\d)([A-ZÇĞİÖŞÜ]{1,3})\b/g, "$1 $2").replace(/(\w)-(\w)/g, "$1 $2");
  s = s.replace(/[ \t]+/g, " ").replace(/\s*\n+\s*/g, ". ").replace(/\.\s*\./g, ".").replace(/\s+([,.!?])/g, "$1").trim();
  if (s.length > max) {
    const kes = s.slice(0, max);
    const son = Math.max(kes.lastIndexOf(". "), kes.lastIndexOf("! "), kes.lastIndexOf("? "));
    s = (son > max * 0.5 ? kes.slice(0, son + 1) : kes).trim();
  }
  return s;
}

/** Tarayıcı sesleri arasından Türkçe olanı seçer: önce erkek/“Google” sesi, yoksa ilk tr sesi. */
export function tarayiciSesiSec(sesler: SpeechSynthesisVoice[]): SpeechSynthesisVoice | null {
  const tr = sesler.filter((v) => /^tr[-_]/i.test(v.lang) || v.lang === "tr");
  if (!tr.length) return null;
  const tercih = tr.find((v) => /ahmet|emel|tolga|erkek|male/i.test(v.name)) || tr.find((v) => /google/i.test(v.name)) || tr.find((v) => !v.localService) || tr[0];
  return tercih;
}

/** Tarayıcının kendi sesiyle okur; bitince (ya da hata/iptal) çözülür. */
export function tarayiciSesi(metin: string): Promise<void> {
  return new Promise((resolve) => {
    if (typeof window === "undefined" || !("speechSynthesis" in window) || !metin) return resolve();
    const synth = window.speechSynthesis;
    const konus = () => {
      const u = new SpeechSynthesisUtterance(metin);
      u.lang = "tr-TR";
      const ses = tarayiciSesiSec(synth.getVoices());
      if (ses) u.voice = ses;
      u.rate = 1.02;
      u.pitch = 0.9; // biraz pes: sakin asistan tonu
      let bitti = false;
      const son = () => { if (!bitti) { bitti = true; resolve(); } };
      u.onend = son;
      u.onerror = son;
      // Bazı tarayıcılar onend'i hiç ateşlemez: uzunluğa göre emniyet zamanı
      setTimeout(son, 4_000 + metin.length * 90);
      synth.cancel();
      synth.speak(u);
    };
    if (synth.getVoices().length) konus();
    else {
      let calisti = false;
      const bir = () => { if (!calisti) { calisti = true; konus(); } };
      synth.addEventListener("voiceschanged", bir, { once: true });
      setTimeout(bir, 600);
    }
  });
}
