import { NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import { getSessionUser } from "@/lib/auth";
import { FRAME_PROFILES } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";

export const runtime = "nodejs";
export const maxDuration = 60;

// Katalog bağlamı — yalnızca toptan LİSTE fiyatları içerir.
function catalogContext(): string {
  const frames = FRAME_PROFILES.map(
    (f) =>
      `${f.code} | ${f.series} Serisi | koli: ${f.koliAdet} adet / ${f.koliMetraj} mt | liste: $${f.priceUSD}/mt | stok: ${f.stok}`
  ).join("\n");
  const tech = TECHNICAL_PRODUCTS.map(
    (t) =>
      `${t.name} | ${t.category} | kutu: ${t.adetPerKutu} | fiyat: ${
        t.priceTL != null ? `₺${t.priceTL}` : `€${t.priceEUR}`
      }`
  ).join("\n");
  return `ÇERÇEVE PROFİLLERİ (toptan liste fiyatları, USD/mt, KDV hariç):\n${frames}\n\nTEKNİK MALZEMELER (kutu fiyatları):\n${tech}`;
}

const BASE_SYSTEM = `Sen Olga Çerçeve Sanayi ve Ticaret Limited Şirketi'nin ürün asistanısın.
Firma bilgileri: Çerçeve profili üretimi/ithalatı, çerçeveleme teknik malzemeleri ve makineleri toptan satışı.
Sipariş Hattı: 0850 305 75 45. Ankara (Merkez): Birlik Mah. 448. Cd. No:56 Çankaya, 0312 495 75 45.
İstanbul: Masko Mobilya Sanayi Sitesi 3-B1 No:4 İkitelli/Başakşehir, 0212 675 27 50.
Web: olgacerceve.com. Çalışma saatleri: Pazartesi–Cumartesi 09:00–18:00.

Kurallar:
- Yalnızca aşağıdaki katalogdaki LİSTE fiyatlarını paylaş. Fiyatlar KDV hariçtir; bunu belirt.
- Alış fiyatı, maliyet, kâr marjı veya iç hesaplama sorulursa kibarca yanıt verme; sadece liste fiyatı paylaşabildiğini söyle.
- Katalogda olmayan ürünler için sipariş hattına yönlendir.
- Türkçe, kısa ve net cevap ver.
- Profiller koli bazında satılır; koli adet/metraj bilgisini gerektiğinde belirt.`;

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 });
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    return NextResponse.json(
      {
        ok: false,
        error:
          "AI asistanı henüz yapılandırılmadı. Vercel ortam değişkenlerine ANTHROPIC_API_KEY ekleyin.",
      },
      { status: 503 }
    );
  }

  const body = await req.json().catch(() => null);
  const messages = Array.isArray(body?.messages) ? body.messages : [];
  const cleaned = messages
    .filter(
      (m: any) =>
        (m?.role === "user" || m?.role === "assistant") &&
        typeof m?.content === "string" &&
        m.content.trim()
    )
    .slice(-20)
    .map((m: any) => ({ role: m.role, content: String(m.content).slice(0, 4000) }));

  if (!cleaned.length || cleaned[cleaned.length - 1].role !== "user") {
    return NextResponse.json({ ok: false, error: "Mesaj gerekli." }, { status: 400 });
  }

  const client = new Anthropic();

  try {
    const response = await client.messages.create({
      model: "claude-opus-5",
      max_tokens: 2048,
      output_config: { effort: "low" },
      system: [
        {
          type: "text",
          text: `${BASE_SYSTEM}\n\n=== KATALOG ===\n${catalogContext()}`,
          cache_control: { type: "ephemeral" },
        },
        {
          type: "text",
          text: `Konuşulan kişi: ${user.name} (${user.role === "staff" ? "firma çalışanı" : "bayi/müşteri"}).`,
        },
      ],
      messages: cleaned,
    });

    if (response.stop_reason === "refusal") {
      return NextResponse.json({
        ok: true,
        reply: "Bu soruya yanıt veremiyorum. Başka nasıl yardımcı olabilirim?",
      });
    }

    const reply = response.content
      .filter((b) => b.type === "text")
      .map((b) => (b as { type: "text"; text: string }).text)
      .join("\n");

    return NextResponse.json({ ok: true, reply });
  } catch (err) {
    console.error("AI hatası:", err);
    return NextResponse.json(
      { ok: false, error: "AI asistanına şu an ulaşılamıyor." },
      { status: 502 }
    );
  }
}
