import { NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import { getSessionUser } from "@/lib/auth";
import { FRAME_PROFILES, findProfile } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { GLASS_TYPES as PLATE_GLASS, GLASS_SIZES, AYNA_SIZES } from "@/data/glass";
import {
  MAT_TYPES,
  GLASS_TYPES as RETAIL_GLASS,
  PRINT_TYPES,
  computeRetailCosts,
} from "@/data/perakende";
import { getDailyRates, istanbulDateKey, listAllOrders, orderBalance } from "@/lib/orders";
import { retailFramePrice } from "@/lib/retail-orders";
import { listCustomers, customerTitle } from "@/lib/customers";
import { getStockData } from "@/lib/stock-store";
import { searchStock, toBoy, BOY_LENGTH } from "@/lib/stock-search";

export const runtime = "nodejs";
export const maxDuration = 60;

// ---------------------------------------------------------------------------
// Statik bağlam — yalnızca toptan LİSTE ve perakende SATIŞ fiyatları içerir.
// Alış fiyatı, maliyet yüzdesi ve çarpanlar hiçbir zaman bağlama girmez.
// ---------------------------------------------------------------------------

function catalogContext(): string {
  const frames = FRAME_PROFILES.map(
    (f) =>
      `${f.code} | ${f.series} Serisi | koli: ${f.koliAdet} adet / ${f.koliMetraj} mt | liste: $${f.priceUSD}/mt`
  ).join("\n");

  const tech = TECHNICAL_PRODUCTS.map(
    (t) =>
      `${t.code} | ${t.name} | ${t.category} | kutu: ${t.adetPerKutu} | ${
        t.priceTL != null ? `₺${t.priceTL}` : `€${t.priceEUR}`
      }`
  ).join("\n");

  const plates = PLATE_GLASS.map((g) => {
    const sizes = (GLASS_SIZES[g.key] || []).map((s) => s.label).join(" · ");
    return `${g.name}: ${sizes}`;
  }).join("\n");
  const ayna = AYNA_SIZES.map((s) => s.label).join(" · ");

  const mats = MAT_TYPES.filter((m) => m.price > 0)
    .map((m) => `${m.name} (${m.code}): ₺${m.price}/m²`)
    .join("\n");
  const rGlass = RETAIL_GLASS.filter((g) => g.price > 0)
    .map((g) => `${g.name}: ₺${g.price}/m² — ${g.desc}`)
    .join("\n");
  const prints = PRINT_TYPES.filter((p) => p.usdPerM2 > 0)
    .map((p) => `${p.name}: $${p.usdPerM2}/m² (KDV dahil, TL = kur × tutar) — ${p.desc}`)
    .join("\n");

  return `=== TOPTAN ÇERÇEVE PROFİLLERİ (liste fiyatı USD/mt, KDV hariç) ===
${frames}

=== TEKNİK MALZEMELER (kutu fiyatları, KDV hariç) ===
${tech}

=== CAM & AYNA PLAKA ÖLÇÜLERİ (toptan, plaka olarak satılır) ===
${plates}
Ayna: ${ayna}

=== PERAKENDE ÇERÇEVELETME MALZEMELERİ (KDV dahil perakende satış) ===
Paspartu / karton:
${mats}
Cam:
${rGlass}
Baskı (eserin kendi alanı üzerinden):
${prints}`;
}

const BASE_SYSTEM = `Sen Olga Çerçeve Sanayi ve Ticaret Limited Şirketi'nin ürün ve sipariş asistanısın.
Firma: Çerçeve profili üretimi/ithalatı, çerçeveleme teknik malzemeleri ve makineleri toptan satışı;
ayrıca perakende çerçeveletme hizmeti.
Sipariş Hattı: 0850 305 75 45 · Web: olgacerceve.com · Çalışma saatleri: Pazartesi–Cumartesi 09:00–18:00.
Ankara (Merkez): Birlik Mah. 448. Cd. No:56 Çankaya/Ankara · 0312 495 75 45.
İstanbul: Tahtakale Mah. Fırat Cd. No:6, Tem34 Sitesi No:95, 34320 Avcılar/İstanbul · 0212 675 27 50.

İŞ KURALLARI:
- Profiller koli bazında satılır. 1 boy = ${BOY_LENGTH} metre. Koli metrajı katalogda yazar.
- Toptan profil fiyatları USD/mt ve KDV hariçtir; TL karşılığı için günün kurunu kullan.
- Teknik malzemeler kutu fiyatıdır; € fiyatlı ürünlerde euro kuru geçerlidir.
- KDV oranı %20'dir. Toptan fiyatlara KDV dahil değildir; perakende malzeme fiyatları KDV dahildir.
- Cam ve ayna toptanda tam plaka satılır, kesim yapılmaz.
- Perakende çerçeveletmede dış ölçü = eser ölçüsü + paspartu kenarları; çevre hesabına 30 cm fire eklenir.

ARAÇ KULLANIMI (çok önemli):
- Stok sorulursa MUTLAKA stok_sorgula aracını çağır. Katalogdaki bilgi statiktir, güncel değildir.
- TL fiyat, kur veya "kaç para tutar" sorulursa MUTLAKA gunun_kuru aracını çağır. Kur uydurma.
- Perakende çerçeveletme fiyatı sorulursa perakende_hesapla aracını kullan, elle hesaplama yapma.
- Sipariş veya müşteri bilgisi sorulursa ilgili arama aracını çağır; hafızadan cevap verme.
- Araç sonucu boş dönerse "kayıt bulunamadı" de; tahmin üretme.

CEVAP KURALLARI:
- Türkçe, kısa ve net yaz. Sayıları binlik ayraçla ve ₺/$/€ işaretiyle ver.
- Yalnızca liste ve perakende SATIŞ fiyatlarını paylaş.
- Alış fiyatı, maliyet, kâr marjı, iskonto tabanı, fiyat çarpanı veya iç hesaplama yöntemi
  sorulursa bu bilgileri paylaşamayacağını söyle; formülü, katsayıyı veya oranı asla açıklama.
- Katalogda olmayan ürünler için sipariş hattına yönlendir.
- Araç adlarını (stok_sorgula, perakende_hesapla gibi) kullanıcıya söyleme; doğal dille anlat.`;

// ---------------------------------------------------------------------------
// Araçlar
// ---------------------------------------------------------------------------

const TOOL_STOK: Anthropic.Tool = {
  name: "stok_sorgula",
  description:
    "Günlük yüklenen stok listesinden bir profil kodunun Ankara ve İstanbul depolarındaki güncel metrajını döndürür. Kod eksik/yanlış yazılmış olabilir, bulanık arama yapar.",
  input_schema: {
    type: "object",
    properties: {
      kod: { type: "string", description: "Profil kodu, örn. GC065 veya gc065-1473" },
    },
    required: ["kod"],
  },
};

const TOOL_KUR: Anthropic.Tool = {
  name: "gunun_kuru",
  description:
    "Bugün için panelde kayıtlı USD ve EUR kurunu döndürür. TL fiyat hesaplamadan önce mutlaka çağır.",
  input_schema: { type: "object", properties: {} },
};

const TOOL_PERAKENDE: Anthropic.Tool = {
  name: "perakende_hesapla",
  description:
    "Perakende çerçeveletme fiyatını hesaplar. Eser ölçüsü, profil kodu ve seçilen malzemelere göre KDV dahil satış tutarını döndürür.",
  input_schema: {
    type: "object",
    properties: {
      profil: { type: "string", description: "Çerçeve profili kodu, örn. GC065" },
      en_cm: { type: "number", description: "Eser genişliği (cm)" },
      boy_cm: { type: "number", description: "Eser yüksekliği (cm)" },
      paspartu_kenar_cm: {
        type: "number",
        description: "Her kenardaki paspartu genişliği (cm). Paspartu yoksa 0.",
      },
      paspartu_tipi: {
        type: "string",
        description: `Paspartu türü: ${MAT_TYPES.map((m) => m.name).join(" / ")}`,
      },
      cam_tipi: {
        type: "string",
        description: `Cam türü: ${RETAIL_GLASS.map((g) => g.name).join(" / ")}`,
      },
      baski_tipi: {
        type: "string",
        description: `Baskı türü: ${PRINT_TYPES.map((p) => p.name).join(" / ")}`,
      },
      adet: { type: "number", description: "Adet (varsayılan 1)" },
    },
    required: ["profil", "en_cm", "boy_cm"],
  },
};

const TOOL_SIPARIS: Anthropic.Tool = {
  name: "siparis_ara",
  description:
    "Kayıtlı toptan siparişlerde arama yapar. Müşteri adı, sipariş numarası veya ürün koduyla arayabilir; boş bırakılırsa son siparişleri döndürür.",
  input_schema: {
    type: "object",
    properties: {
      sorgu: { type: "string", description: "Müşteri adı, sipariş no veya ürün kodu" },
      gun: { type: "number", description: "Son kaç günün siparişleri (varsayılan 90)" },
      limit: { type: "number", description: "En fazla kaç kayıt (varsayılan 15)" },
    },
  },
};

const TOOL_MUSTERI: Anthropic.Tool = {
  name: "musteri_ara",
  description:
    "Müşteri defterinde arama yapar. Firma/kişi adı, telefon veya şehirle arayabilir. Cari bakiye ve sipariş özeti de döner.",
  input_schema: {
    type: "object",
    properties: {
      sorgu: { type: "string", description: "Firma adı, kişi adı, telefon veya şehir" },
      limit: { type: "number", description: "En fazla kaç kayıt (varsayılan 10)" },
    },
    required: ["sorgu"],
  },
};

const trNorm = (s: string) =>
  String(s || "")
    .toLocaleLowerCase("tr-TR")
    .replace(/[çğıöşü]/g, (c) => ({ ç: "c", ğ: "g", ı: "i", ö: "o", ş: "s", ü: "u" }[c] || c))
    .replace(/\s+/g, " ")
    .trim();

async function runTool(name: string, input: any): Promise<string> {
  switch (name) {
    case "stok_sorgula": {
      const kod = String(input?.kod || "").trim();
      const data = await getStockData();
      const matches = searchStock(data.items, kod, 0.72, 12);
      if (!matches.length) {
        return JSON.stringify({
          bulunamadi: true,
          not: `"${kod}" stok listesinde bulunamadı.`,
          stokGuncelleme: data.updatedAt,
        });
      }
      return JSON.stringify({
        stokGuncelleme: data.updatedAt,
        kaynak: data.sourceName,
        sonuclar: matches.map((m) => ({
          kod: m.item.code,
          ankaraMt: Math.round(m.item.ankaraMt * 10) / 10,
          ankaraBoy: toBoy(m.item.ankaraMt),
          istanbulMt: Math.round(m.item.istanbulMt * 10) / 10,
          istanbulBoy: toBoy(m.item.istanbulMt),
          toplamMt: Math.round((m.item.ankaraMt + m.item.istanbulMt) * 10) / 10,
        })),
      });
    }

    case "gunun_kuru": {
      const key = istanbulDateKey();
      const rates = await getDailyRates(key);
      if (!rates || (!rates.rate && !rates.euroRate)) {
        return JSON.stringify({
          bulunamadi: true,
          not: "Bugün için kur girilmemiş. Sipariş panelinden günün kuru girilmeli.",
        });
      }
      return JSON.stringify({
        tarih: key,
        usd: rates.rate || null,
        eur: rates.euroRate || null,
        guncelleyen: rates.by || null,
      });
    }

    case "perakende_hesapla": {
      const kod = String(input?.profil || "").trim();
      const profile = findProfile(kod);
      if (!profile) {
        return JSON.stringify({ bulunamadi: true, not: `"${kod}" katalogda bulunamadı.` });
      }
      const rates = await getDailyRates(istanbulDateKey());
      const usd = rates?.rate || 0;
      if (!(usd > 0)) {
        return JSON.stringify({
          bulunamadi: true,
          not: "Bugünün USD kuru girilmediği için perakende fiyat hesaplanamıyor.",
        });
      }
      const fp = retailFramePrice(profile.code, usd);
      const kenarMM = Math.max(0, Number(input?.paspartu_kenar_cm) || 0) * 10;
      const mat =
        MAT_TYPES.find((m) => trNorm(m.name) === trNorm(input?.paspartu_tipi)) ||
        (kenarMM > 0 ? MAT_TYPES[1] : MAT_TYPES[0]);
      const glass =
        RETAIL_GLASS.find((g) => trNorm(g.name) === trNorm(input?.cam_tipi)) || RETAIL_GLASS[1];
      const print =
        PRINT_TYPES.find((p) => trNorm(p.name) === trNorm(input?.baski_tipi)) || PRINT_TYPES[0];
      const adet = Math.max(1, Math.round(Number(input?.adet) || 1));

      const costs = computeRetailCosts({
        wMM: Math.max(0, Number(input?.en_cm) || 0) * 10,
        hMM: Math.max(0, Number(input?.boy_cm) || 0) * 10,
        matTop: kenarMM,
        matRight: kenarMM,
        matBottom: kenarMM,
        matLeft: kenarMM,
        framePriceTL: fp.tlPerM,
        matPrice: mat.price,
        doubleMat: false,
        innerMatPrice: 0,
        zeminEnabled: false,
        zeminPrice: 0,
        glassPrice: glass.price,
        printUsdPerM2: print.usdPerM2,
        usdRate: usd,
      });

      const r2 = (n: number) => Math.round(n * 100) / 100;
      return JSON.stringify({
        profil: profile.code,
        eser: `${input?.en_cm} × ${input?.boy_cm} cm`,
        paspartu: mat.price > 0 ? `${mat.name} (${input?.paspartu_kenar_cm || 0} cm kenar)` : "Yok",
        cam: glass.name,
        baski: print.name,
        adet,
        kalemler: {
          cerceveTL: r2(costs.frameCost),
          paspartuTL: r2(costs.matCost),
          camTL: r2(costs.glassCost),
          baskiTL: r2(costs.printCost),
        },
        birimTL: r2(costs.itemTotal),
        toplamTL: r2(costs.itemTotal * adet),
        not: "Perakende satış fiyatı, KDV dahil.",
      });
    }

    case "siparis_ara": {
      const q = trNorm(input?.sorgu || "");
      const gun = Math.min(730, Math.max(1, Number(input?.gun) || 90));
      const limit = Math.min(40, Math.max(1, Number(input?.limit) || 15));
      const since = new Date(Date.now() - gun * 86400000).toISOString().slice(0, 10);

      const all = await listAllOrders();
      const filtered = all
        .filter((o) => o.dateKey >= since)
        .filter((o) => {
          if (!q) return true;
          if (trNorm(o.customer).includes(q)) return true;
          if (trNorm(o.orderId).includes(q)) return true;
          if (trNorm(o.employee).includes(q)) return true;
          return (o.lines || []).some((l: any) => trNorm(l?.code || l?.name || "").includes(q));
        })
        .slice(0, limit);

      if (!filtered.length) {
        return JSON.stringify({ bulunamadi: true, not: "Bu kritere uyan sipariş yok.", toplamKayit: all.length });
      }
      return JSON.stringify({
        bulunan: filtered.length,
        siparisler: filtered.map((o) => ({
          no: o.orderId,
          tarih: o.dateKey,
          musteri: o.customer,
          alan: o.employee,
          durum: o.status,
          odeme: o.payment || "bekliyor",
          netTL: Math.round(o.net),
          kalanTL: Math.round(orderBalance(o)),
          satirlar: (o.lines || [])
            .slice(0, 12)
            .map((l: any) => `${l?.code || l?.name || "?"} × ${l?.qty ?? ""} ${l?.unit ?? ""}`.trim()),
          not: o.note || undefined,
        })),
      });
    }

    case "musteri_ara": {
      const q = trNorm(input?.sorgu || "");
      const limit = Math.min(25, Math.max(1, Number(input?.limit) || 10));
      const [customers, orders] = await Promise.all([listCustomers(), listAllOrders()]);
      const hits = customers
        .filter((c) => {
          const hay = trNorm(
            [customerTitle(c), c.company, c.firstName, c.lastName, c.phone, c.city, c.district].join(" ")
          );
          return !q || hay.includes(q);
        })
        .slice(0, limit);

      if (!hits.length) {
        return JSON.stringify({ bulunamadi: true, not: "Müşteri defterinde eşleşme yok.", toplamKayit: customers.length });
      }
      return JSON.stringify({
        bulunan: hits.length,
        musteriler: hits.map((c) => {
          const mine = orders.filter(
            (o) => o.customerId === c.id || trNorm(o.customer) === trNorm(customerTitle(c))
          );
          const bakiye = mine.reduce((s, o) => s + orderBalance(o), 0);
          return {
            ad: customerTitle(c),
            telefon: c.phone || undefined,
            sehir: [c.district, c.city].filter(Boolean).join(" / ") || undefined,
            sube: c.branch,
            siparisAdedi: mine.length,
            sonSiparis: mine[0] ? `${mine[0].orderId} · ${mine[0].dateKey}` : undefined,
            acikBakiyeTL: Math.round(bakiye),
            not: c.note || undefined,
          };
        }),
      });
    }

    default:
      return JSON.stringify({ hata: "Bilinmeyen araç." });
  }
}

// ---------------------------------------------------------------------------

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
  const incoming = Array.isArray(body?.messages) ? body.messages : [];
  const cleaned: Anthropic.MessageParam[] = incoming
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

  const isStaff = user.role === "staff";
  // Sipariş ve müşteri kayıtları yalnızca firma çalışanlarına açıktır.
  const tools: Anthropic.Tool[] = isStaff
    ? [TOOL_STOK, TOOL_KUR, TOOL_PERAKENDE, TOOL_SIPARIS, TOOL_MUSTERI]
    : [TOOL_STOK, TOOL_KUR, TOOL_PERAKENDE];

  const client = new Anthropic();
  const messages: Anthropic.MessageParam[] = [...cleaned];

  try {
    for (let round = 0; round < 6; round++) {
      const response = await client.messages.create({
        model: "claude-opus-5",
        max_tokens: 2048,
        output_config: { effort: "medium" },
        system: [
          {
            type: "text",
            text: `${BASE_SYSTEM}\n\n${catalogContext()}`,
            cache_control: { type: "ephemeral" },
          },
          {
            type: "text",
            text: `Bugün: ${istanbulDateKey()}. Konuşulan kişi: ${user.name} (${
              isStaff ? "firma çalışanı — sipariş ve müşteri kayıtlarını görebilir" : "bayi/müşteri — yalnızca ürün, stok ve fiyat bilgisi alabilir"
            }).`,
          },
        ],
        tools,
        messages,
      });

      if (response.stop_reason === "refusal") {
        return NextResponse.json({
          ok: true,
          reply: "Bu soruya yanıt veremiyorum. Başka nasıl yardımcı olabilirim?",
        });
      }

      const toolUses = response.content.filter(
        (b): b is Anthropic.ToolUseBlock => b.type === "tool_use"
      );

      if (!toolUses.length) {
        const reply = response.content
          .filter((b) => b.type === "text")
          .map((b) => (b as { type: "text"; text: string }).text)
          .join("\n");
        return NextResponse.json({ ok: true, reply });
      }

      messages.push({ role: "assistant", content: response.content });
      const results: Anthropic.ToolResultBlockParam[] = await Promise.all(
        toolUses.map(async (t) => {
          let content: string;
          try {
            content = await runTool(t.name, t.input);
          } catch (err) {
            console.error("Araç hatası:", t.name, err);
            content = JSON.stringify({ hata: "Veri okunamadı." });
          }
          return { type: "tool_result" as const, tool_use_id: t.id, content };
        })
      );
      messages.push({ role: "user", content: results });
    }

    return NextResponse.json({
      ok: true,
      reply: "Sorguyu tamamlayamadım. Daha dar bir soru sorar mısınız?",
    });
  } catch (err) {
    console.error("AI hatası:", err);
    return NextResponse.json(
      { ok: false, error: "AI asistanına şu an ulaşılamıyor." },
      { status: 502 }
    );
  }
}
