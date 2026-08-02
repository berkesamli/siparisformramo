import { NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import { getSessionUser } from "@/lib/auth";
import { FRAME_PROFILES, findProfile } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { GLASS_TYPES } from "@/data/glass";

export const runtime = "nodejs";
export const maxDuration = 60;

// WhatsApp/telefon notundan gelen serbest sipariş metnini forma
// dökülebilecek satırlara çevirir. Fiyat üretmez — fiyatlar katalogdan
// ve günün kurundan formda hesaplanır.

const SYSTEM = `Olga Çerçeve'nin sipariş metni çözümleyicisisin. Sana müşteriden gelen
serbest yazılmış (WhatsApp, telefon notu) sipariş metni verilir. Görevin bunu
yapılandırılmış sipariş satırlarına çevirmek.

KURALLAR:
- Sadece metinde YAZAN ürünleri çıkar. Uydurma, tahmini ürün ekleme.
- Ürün kodları eksik/yanlış yazılmış olabilir: "gc065 1473", "ks2030", "GB 139-1211t"
  gibi. Katalogdaki en yakın kodu bul ve "code" alanına KATALOGDAKİ tam kodu yaz.
  Emin değilsen kullanıcının yazdığını olduğu gibi bırak ve confidence'ı düşür.
- kind alanı: çerçeve profili → "frame", cam → "glass", ayna → "ayna",
  teknik malzeme (agraf, askı, bant, vida, çivi vb.) → "technical",
  katalogda karşılığı olmayan diğer her şey → "other".
- Çerçeve birimi (unit): metre / boy / koli. Metinde "koli" geçiyorsa "koli",
  "boy" geçiyorsa "boy", aksi halde "metre". Bir boy 2,9 metredir.
- Miktar sayısal olmalı. "3 koli" → qty 3, unit koli. "50 mt" → qty 50, unit metre.
- Renk kodu ürün kodunun parçasıysa koda dahil et (GB139-1211T gibi).
  Ayrı bir açıklamaysa note alanına yaz.
- Cam/ayna için ölçü metinde geçiyorsa note'a yaz.
- confidence: 1 = kod katalogda birebir bulundu, 0.7 = büyük olasılıkla doğru,
  0.4 = tahmin, kullanıcı kontrol etmeli.
- Müşteri adı, teslimat notu gibi ürün olmayan bilgileri "customer" ve "note"
  alanlarına ayır; satır olarak ekleme.`;

const TOOL = {
  name: "siparis_satirlari",
  description: "Çözümlenen sipariş satırlarını döndürür.",
  input_schema: {
    type: "object" as const,
    properties: {
      customer: {
        type: "string",
        description: "Metinde geçen müşteri/firma adı, yoksa boş bırak.",
      },
      note: {
        type: "string",
        description: "Teslimat, kargo, aciliyet gibi genel notlar.",
      },
      lines: {
        type: "array",
        items: {
          type: "object",
          properties: {
            kind: {
              type: "string",
              enum: ["frame", "glass", "ayna", "technical", "other"],
            },
            code: { type: "string", description: "Ürün/profil kodu veya ürün adı" },
            unit: { type: "string", enum: ["metre", "boy", "koli", "adet", "kutu"] },
            qty: { type: "number" },
            note: { type: "string" },
            confidence: { type: "number" },
          },
          required: ["kind", "code", "qty"],
        },
      },
    },
    required: ["lines"],
  },
};

function catalogList(): string {
  const frames = FRAME_PROFILES.map(
    (f) => `${f.code} (${f.series} serisi, koli ${f.koliAdet} adet / ${f.koliMetraj} mt)`
  ).join("\n");
  const tech = TECHNICAL_PRODUCTS.map((t) => `${t.name} (${t.category})`).join("\n");
  const glass = GLASS_TYPES.map((g) => g.name).join(", ");
  return `ÇERÇEVE PROFİLLERİ:\n${frames}\n\nTEKNİK MALZEMELER:\n${tech}\n\nCAM TÜRLERİ: ${glass}\nAYNA: 2mm / 3mm / 4mm plaka`;
}

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    return NextResponse.json(
      {
        ok: false,
        error:
          "Yapay zeka henüz yapılandırılmadı. Vercel ortam değişkenlerine ANTHROPIC_API_KEY ekleyin.",
      },
      { status: 503 }
    );
  }

  const body = await req.json().catch(() => null);
  const text = String(body?.text || "").trim().slice(0, 6000);
  if (!text) {
    return NextResponse.json({ ok: false, error: "Metin gerekli." }, { status: 400 });
  }

  const client = new Anthropic();

  try {
    const response = await client.messages.create({
      model: "claude-opus-5",
      max_tokens: 4096,
      output_config: { effort: "low" },
      system: [
        {
          type: "text",
          text: `${SYSTEM}\n\n=== KATALOG ===\n${catalogList()}`,
          cache_control: { type: "ephemeral" },
        },
      ],
      tools: [TOOL],
      tool_choice: { type: "tool", name: "siparis_satirlari" },
      messages: [{ role: "user", content: text }],
    });

    const toolUse = response.content.find(
      (c): c is Anthropic.ToolUseBlock => c.type === "tool_use"
    );
    if (!toolUse) {
      return NextResponse.json(
        { ok: false, error: "Metin çözümlenemedi, elle girmeyi deneyin." },
        { status: 502 }
      );
    }

    const parsed = toolUse.input as {
      customer?: string;
      note?: string;
      lines?: any[];
    };

    // Kodları katalogla doğrula: bulunanı tam koda çevir, bulunamayanı işaretle
    const lines = (parsed.lines || []).slice(0, 60).map((l: any) => {
      const rawCode = String(l?.code || "").trim();
      const kind = ["frame", "glass", "ayna", "technical", "other"].includes(l?.kind)
        ? l.kind
        : "other";
      let code = rawCode;
      let matched = false;
      if (kind === "frame") {
        const p = findProfile(rawCode);
        if (p) {
          code = p.code;
          matched = true;
        }
      }
      return {
        kind,
        code,
        rawCode,
        matched,
        unit: String(l?.unit || "metre"),
        qty: Math.max(0, Number(l?.qty) || 0),
        note: String(l?.note || "").slice(0, 200),
        confidence: Math.min(1, Math.max(0, Number(l?.confidence) || 0.5)),
      };
    });

    return NextResponse.json({
      ok: true,
      customer: String(parsed.customer || "").slice(0, 160),
      note: String(parsed.note || "").slice(0, 400),
      lines,
    });
  } catch (err: any) {
    console.error("Sipariş metni çözümlenemedi:", err);
    return NextResponse.json(
      { ok: false, error: "Çözümleme sırasında hata oluştu: " + (err?.message || "bilinmiyor") },
      { status: 500 }
    );
  }
}
