import { NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import { getSessionUser } from "@/lib/auth";
import { FRAME_PROFILES } from "@/data/catalog";
import { siparisMetniCoz, cerceveKodu, teknikUrun, type MetinSatir } from "@/lib/siparis-metin";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { GLASS_TYPES } from "@/data/glass";

export const runtime = "nodejs";
export const maxDuration = 60;

// WhatsApp/telefon notundan gelen serbest sipariş metnini forma
// dökülebilecek satırlara çevirir. Fiyat üretmez — fiyatlar katalogdan
// ve günün kurundan formda hesaplanır.
//
// Önce kural tabanlı çözümleyici (lib/siparis-metin) çalışır: kod + desen/renk
// eki, miktar, birim, iskonto/KDV başlıkları kesin okunur. Yalnızca orada
// okunamayan satırlar yapay zekâya gider; anahtar yoksa bunlar "okunamayan"
// olarak kullanıcıya gösterilir.

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
  const body = await req.json().catch(() => null);
  const text = String(body?.text || "").trim().slice(0, 6000);
  if (!text) {
    return NextResponse.json({ ok: false, error: "Metin gerekli." }, { status: 400 });
  }

  // 1) Kural tabanlı: kesin okunan satırlar
  const det = siparisMetniCoz(text, FRAME_PROFILES, TECHNICAL_PRODUCTS);
  const lines: MetinSatir[] = [...det.lines];
  let customer = det.customer;
  let note = det.note;
  let okunamayan: string[] = [];

  // 2) Kalan satırlar yapay zekâya (anahtar varsa)
  if (det.kalan.length && process.env.ANTHROPIC_API_KEY) {
    try {
      const client = new Anthropic();
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
        messages: [{ role: "user", content: det.kalan.join("\n") }],
      });
      const toolUse = response.content.find(
        (c): c is Anthropic.ToolUseBlock => c.type === "tool_use"
      );
      const parsed = (toolUse?.input || {}) as { customer?: string; note?: string; lines?: any[] };
      for (const l of (parsed.lines || []).slice(0, 60)) {
        const rawCode = String(l?.code || "").trim();
        const kind = ["frame", "glass", "ayna", "technical", "other"].includes(l?.kind) ? (l.kind as MetinSatir["kind"]) : "other";
        const qty = Math.max(0, Number(l?.qty) || 0);
        const unit = String(l?.unit || "metre");
        const lnote = String(l?.note || "").slice(0, 200);
        const conf = Math.min(1, Math.max(0, Number(l?.confidence) || 0.5));
        if (kind === "frame") {
          // Desen/renk eki korunur; ana kod katalogdan, stok listesindeki yazımla
          const c = cerceveKodu(rawCode, FRAME_PROFILES);
          lines.push({ kind, code: c.code, rawCode, matched: c.matched, unit, qty, note: [lnote, c.note].filter(Boolean).join(" · "), confidence: c.matched ? Math.max(conf, 0.7) : Math.min(conf, 0.5) });
        } else if (kind === "technical") {
          const t = teknikUrun(rawCode, TECHNICAL_PRODUCTS);
          lines.push({ kind, code: t ? t.t.name + (t.kartonKodu ? ` (${t.kartonKodu})` : "") : rawCode, rawCode, matched: !!t, unit, qty, note: lnote, confidence: t ? Math.max(conf, 0.7) : conf, techCode: t?.t.code, kartonKodu: t?.kartonKodu });
        } else {
          lines.push({ kind, code: rawCode, rawCode, matched: false, unit, qty, note: lnote, confidence: conf });
        }
      }
      if (!customer && parsed.customer) customer = String(parsed.customer).slice(0, 160);
      if (parsed.note) note = [note, String(parsed.note).slice(0, 400)].filter(Boolean).join(" · ");
    } catch (err: any) {
      console.error("Sipariş metni yapay zekâ ile çözümlenemedi:", err);
      okunamayan = det.kalan;
    }
  } else if (det.kalan.length) {
    okunamayan = det.kalan;
  }

  if (!lines.length && !okunamayan.length) {
    return NextResponse.json({ ok: false, error: "Metinde ürün satırı bulunamadı." }, { status: 422 });
  }
  return NextResponse.json({
    ok: true,
    customer: customer.slice(0, 160),
    note: note.slice(0, 400),
    iskontoPct: det.iskontoPct,
    kdv: det.kdv,
    lines: lines.slice(0, 80),
    okunamayan,
    yapayZeka: Boolean(process.env.ANTHROPIC_API_KEY),
  });
}
