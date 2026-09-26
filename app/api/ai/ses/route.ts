// Jarvis'in sesi — metni sese çevirir. ELEVENLABS_API_KEY tanımlıysa ElevenLabs (doğal, Türkçe
// konuşabilen sesler; ELEVENLABS_VOICE_ID ile ses seçilir), yoksa 501 döner ve arayüz tarayıcının
// kendi Türkçe sesine düşer. Yalnızca oturumu olan kullanıcılar; metin 1000 karakterle sınırlı.
import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

// ElevenLabs hazır "Daniel" sesi (İngiliz, sakin) — Jarvis'e en yakın hazır ses; tr-TR'yi çok dilli modelle konuşur
const VARSAYILAN_SES = "onwK4e9ZLuTAKqWW03F9";

function ayar() {
  const anahtar = process.env.ELEVENLABS_API_KEY || "";
  return {
    anahtar,
    ses: process.env.ELEVENLABS_VOICE_ID || VARSAYILAN_SES,
    model: process.env.ELEVENLABS_MODEL || "eleven_multilingual_v2",
  };
}

export async function GET() {
  const user = await getSessionUser();
  if (!user) return NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 });
  const a = ayar();
  return NextResponse.json({ ok: true, saglayici: a.anahtar ? "elevenlabs" : "tarayici", ses: a.anahtar ? a.ses : null });
}

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user) return NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 });
  const a = ayar();
  if (!a.anahtar) return NextResponse.json({ ok: false, error: "Sunucu sesi tanımlı değil (ELEVENLABS_API_KEY); tarayıcı sesi kullanılır." }, { status: 501 });
  const body = (await req.json().catch(() => null)) as { metin?: string } | null;
  const metin = String(body?.metin || "").trim().slice(0, 1000);
  if (!metin) return NextResponse.json({ ok: false, error: "Metin gerekli." }, { status: 400 });
  try {
    const r = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${encodeURIComponent(a.ses)}?output_format=mp3_44100_64`, {
      method: "POST",
      headers: { "xi-api-key": a.anahtar, "Content-Type": "application/json", Accept: "audio/mpeg" },
      body: JSON.stringify({
        text: metin,
        model_id: a.model,
        voice_settings: { stability: 0.55, similarity_boost: 0.8, style: 0.15, use_speaker_boost: true },
      }),
      signal: AbortSignal.timeout(25_000),
    });
    if (!r.ok || !r.body) {
      const detay = (await r.text().catch(() => "")).slice(0, 300);
      console.error("ElevenLabs hatası:", r.status, detay);
      return NextResponse.json({ ok: false, error: `Ses servisi yanıt vermedi (HTTP ${r.status}).` }, { status: 502 });
    }
    return new Response(r.body, { headers: { "Content-Type": "audio/mpeg", "Cache-Control": "no-store" } });
  } catch (e) {
    console.error("ElevenLabs bağlantı hatası:", (e as Error)?.message);
    return NextResponse.json({ ok: false, error: "Ses servisine ulaşılamadı." }, { status: 502 });
  }
}
