// Yapay zekâ YANIT TASLAĞI — yalnızca taslak üretir, hiçbir şey göndermez.
// Çalışan taslağı okur, düzeltir ve kendisi gönderir. Bağlama yalnızca
// toptan LİSTE ve perakende SATIŞ fiyatları, stok ve müşteri kartı girer;
// alış fiyatı / maliyet asla girmez.

import Anthropic from "@anthropic-ai/sdk";
import { findProfile } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { MAT_TYPES, GLASS_TYPES, PRINT_TYPES } from "@/data/perakende";
import { getDailyRates, istanbulDateKey } from "@/lib/orders";
import { getStockData } from "@/lib/stock-store";
import { searchStock, toBoy } from "@/lib/stock-search";
import { getCustomer, customerTitle, musteriBolgesi, bolgeler } from "@/lib/customers";
import { getRetailCustomer } from "@/lib/retail-customers";
import { cariOzet, mikroConfigured } from "@/lib/mikro";
import { konusma, mesajlar } from "./db";
import { KANAL_ADI, type Konusma, type Mesaj } from "./tur";

export function taslakHazir(): boolean {
  return Boolean((process.env.ANTHROPIC_API_KEY || "").trim());
}

const fmt = (n: number) => (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 2 });

/** Metindeki olası profil kodları (GB211-4110B, KS 3420 BLACK, 3127S-A79 …). */
export function kodAdaylari(metin: string): string[] {
  const out = new Set<string>();
  // Ön ek bitişikse her harf ("gb211"); boşlukluysa BÜYÜK harf ("KS 3420") ya da en çok 2 küçük harf
  // ("ks 3420") — "ile 3127" gibi sözcükler ön ek sayılmaz. Yalın sayı (adet, yıl, fiyat) kod değildir.
  const re = /\b(?:([A-Za-z]{1,3})|([A-Z]{1,3}|[a-z]{1,2})\s)?(\d{3,4})([A-Za-z]{0,2})(?:\s?-\s?([A-Za-z0-9]{1,7}))?\b/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(metin)) && out.size < 8) {
    const onEk = (m[1] || m[2] || "").toUpperCase();
    if (!onEk && !m[4] && !m[5]) continue;
    if (m[2] && /^(19|20)\d\d$/.test(m[3]) && !m[4] && !m[5]) continue; // "tl 2026" → yıl
    const taban = `${onEk}${m[3]}${m[4].toUpperCase()}`;
    out.add(m[5] ? `${taban}-${m[5].toUpperCase()}` : taban);
  }
  return [...out];
}

async function urunBaglami(metinler: string[]): Promise<string> {
  const kodlar = kodAdaylari(metinler.join("\n"));
  if (!kodlar.length) return "";
  let stok: Awaited<ReturnType<typeof getStockData>> | null = null;
  try { stok = await getStockData(); } catch { stok = null; }
  const satirlar: string[] = [];
  for (const kod of kodlar) {
    const taban = kod.split("-")[0];
    const p = findProfile(taban) || findProfile(kod);
    const parca: string[] = [];
    if (p) parca.push(`katalog: ${p.code} ${p.series} serisi, liste $${p.priceUSD}/mt (KDV hariç), koli ${p.koliAdet} boy / ${p.koliMetraj} mt`);
    if (stok?.items?.length) {
      const es = searchStock(stok.items, kod, 0.85, 4);
      if (es.length) {
        parca.push("stok: " + es.map((e) => `${e.item.code} Ankara ${fmt(e.item.ankaraMt)} mt (${toBoy(e.item.ankaraMt)} boy) / İstanbul ${fmt(e.item.istanbulMt)} mt (${toBoy(e.item.istanbulMt)} boy)`).join("; "));
      } else parca.push("stok: listede yok");
    }
    const teknik = TECHNICAL_PRODUCTS.find((t) => t.code.toUpperCase() === kod || t.code.toUpperCase() === taban);
    if (teknik) parca.push(`teknik ürün: ${teknik.name} — kutu ${teknik.adetPerKutu} adet, ${teknik.priceTL != null ? `₺${teknik.priceTL}` : `€${teknik.priceEUR}`} (KDV hariç)`);
    if (parca.length) satirlar.push(`${kod} → ${parca.join(" | ")}`);
  }
  return satirlar.length ? `=== KONUŞMADA GEÇEN ÜRÜNLER ===\n${satirlar.join("\n")}` : "";
}

function perakendeBaglami(): string {
  const mats = MAT_TYPES.filter((m) => m.price > 0).map((m) => `${m.name}: ₺${m.price}/m²`).join(" · ");
  const cam = GLASS_TYPES.filter((g) => g.price > 0).map((g) => `${g.name}: ₺${g.price}/m²`).join(" · ");
  const baski = PRINT_TYPES.filter((p) => p.usdPerM2 > 0).map((p) => `${p.name}: $${p.usdPerM2}/m²`).join(" · ");
  return `=== PERAKENDE ÇERÇEVELETME (KDV dahil satış fiyatları, m² üzerinden) ===\nPaspartu: ${mats}\nCam: ${cam}\nBaskı: ${baski}\nÇerçeve profili perakendede metre fiyatıyla eklenir; kesin fiyat ölçü (en × boy) ve seçilen profile göre hesaplanır.`;
}

async function musteriBaglami(k: Konusma): Promise<string> {
  if (!k.musteriId) return `Karşı taraf müşteri defterinde kayıtlı değil (${k.kanal === "email" ? "e-posta" : "numara/hesap"}: ${k.disKimlik}).`;
  if (k.musteriTur === "perakende") {
    const c = await getRetailCustomer(k.musteriId).catch(() => null);
    return c ? `Perakende müşteri kartı: ${c.name}${c.phone ? ", tel " + c.phone : ""}${c.address ? ", adres " + c.address : ""}${c.note ? ", not: " + c.note : ""}.` : "Perakende müşteri (kart okunamadı).";
  }
  const c = await getCustomer(k.musteriId).catch(() => null);
  if (!c) return "Toptan müşteri (kart okunamadı).";
  const b = musteriBolgesi(c);
  const parca = [
    `Toptan müşteri (bayi): ${customerTitle(c)}`,
    c.city ? `${c.city}${c.district ? "/" + c.district : ""}` : "",
    `bölge ${bolgeler()[b]?.label || b}`,
    c.iskontoPct ? `bayi iskontosu %${c.iskontoPct}` : "",
    c.note ? `not: ${c.note}` : "",
  ].filter(Boolean);
  let mikro = "";
  if (c.mikroCariKod && mikroConfigured()) {
    const r = await cariOzet(c.mikroCariKod).catch(() => null);
    if (r?.ok && r.ozet) mikro = `\nMikro cari bakiyesi: ₺${fmt(r.ozet.bakiye)} (${r.ozet.bakiye > 0 ? "müşteri borçlu" : r.ozet.bakiye < 0 ? "müşteri alacaklı" : "sıfır"}). Bu bilgiyi YALNIZCA müşteri hesabını/borcunu sorduysa, "kayıtlarımıza göre" diyerek ver.`;
  }
  return parca.join(", ") + "." + mikro;
}

function kanalUslubu(k: Konusma): string {
  if (k.kanal === "email") {
    return `Kanal: e-posta. "Merhaba" ya da "Sayın …" ile başla, kısa paragraflar yaz, "Saygılarımızla,\nOlga Çerçeve" ile bitir. Konu: ${k.baslik || "-"}.`;
  }
  return `Kanal: ${KANAL_ADI[k.kanal]} sohbeti. Kısa yaz (1–5 cümle), samimi ama nazik, biçimlendirme (madde işareti, yıldız) kullanma. Gerekiyorsa tek bir soru sor. Sonuna "Olga Çerçeve" ekleme; sohbet dilinde kal.`;
}

function konusmaMetni(k: Konusma, ms: Mesaj[]): string {
  return ms
    .slice(-14)
    .map((m) => {
      const kim = m.yon === "gelen" ? `MÜŞTERİ (${m.gonderen || k.ad})` : `BİZ (${m.gonderen || "çalışan"})`;
      const ek = m.ekler?.length ? ` [ek: ${m.ekler.map((e) => e.ad).join(", ")}]` : "";
      return `${kim} — ${new Date(m.at).toLocaleString("tr-TR", { timeZone: "Europe/Istanbul" })}:\n${(m.govde || "").slice(0, 2500)}${ek}`;
    })
    .join("\n\n");
}

const SISTEM = `Sen Olga Çerçeve Sanayi ve Ticaret Limited Şirketi'nin müşteri iletişim asistanısın; firma çalışanı adına GÖNDERİLMEYECEK, çalışanın gözden geçireceği bir YANIT TASLAĞI yazıyorsun.
Firma: çerçeve profili üretimi/ithalatı, çerçeveleme teknik malzemeleri ve makineleri toptan satışı; ayrıca perakende çerçeveletme hizmeti. Ankara ve İstanbul'da depo/şube. Sipariş hattı 0850 305 75 45, web olgacerceve.com, Pazartesi–Cumartesi 09:00–18:00.
Kurallar:
- Türkçe yaz, "siz" diye hitap et. Müşterinin son mesajına doğrudan cevap ver; gereksiz uzatma.
- Yalnızca aşağıdaki bağlamda verilen fiyat, stok ve kayıt bilgisini kullan. Bağlamda olmayan bir fiyatı/ölçüyü/teslim tarihini ASLA uydurma; bilmiyorsan "kontrol edip size döneceğiz" de.
- Toptan (bayi) fiyatları USD/metre ve KDV hariçtir; perakende fiyatlar TL ve KDV dahildir. Toptan liste fiyatı verirken günün kurunu belirt, TL karşılığını yaklaşık olarak ver.
- Stok bilgisini "şu an" diyerek ver; kesin rezervasyon sözü verme.
- Mikro bakiyesi verilmişse yalnızca müşteri hesabını sorduysa söyle.
- Müşteri sipariş vermek istiyorsa ürün kodu, adet/metre ve teslimat şubesi (Ankara/İstanbul) gibi eksikleri tek seferde sor.
- Şikâyet/olumsuz durumda önce anlayış göster, sonra çözüm adımı yaz.
- Yalnızca yanıt metnini döndür; başlık, açıklama, tırnak ya da "Taslak:" gibi ekler yazma.`;

export async function taslakUret(konusmaId: string): Promise<{ taslak: string; model: string }> {
  if (!taslakHazir()) throw new Error("Yapay zekâ anahtarı (ANTHROPIC_API_KEY) tanımlı değil.");
  const k = await konusma(konusmaId);
  if (!k) throw new Error("Konuşma bulunamadı.");
  const ms = await mesajlar(k.id, 40);
  if (!ms.length) throw new Error("Konuşmada mesaj yok.");
  const gelenMetinler = ms.filter((m) => m.yon === "gelen").slice(-6).map((m) => m.govde);
  const [urun, musteri, kur] = await Promise.all([
    urunBaglami(gelenMetinler),
    musteriBaglami(k),
    getDailyRates(istanbulDateKey()).catch(() => null),
  ]);
  const baglam = [
    `Bugün: ${istanbulDateKey()}.${kur ? ` Günün kuru: ${fmt(kur.rate)} TL/USD${kur.euroRate ? `, ${fmt(kur.euroRate)} TL/EUR` : ""}.` : " Günün kuru girilmemiş; TL karşılığı verme."}`,
    kanalUslubu(k),
    `Karşı taraf: ${k.ad || k.disKimlik}. ${musteri}`,
    urun,
    perakendeBaglami(),
  ].filter(Boolean).join("\n\n");

  const client = new Anthropic();
  const model = "claude-opus-5";
  const r = await client.messages.create({
    model,
    max_tokens: 900,
    output_config: { effort: "low" },
    system: [{ type: "text", text: SISTEM }, { type: "text", text: baglam }],
    messages: [{ role: "user", content: `KONUŞMA:\n\n${konusmaMetni(k, ms)}\n\nMüşterinin son mesajına firma adına gönderilecek yanıt taslağını yaz.` }],
  });
  const metin = r.content.filter((c) => c.type === "text").map((c) => (c as { text: string }).text).join("\n").trim();
  if (!metin) throw new Error("Taslak üretilemedi.");
  return { taslak: metin.replace(/^["“]|["”]$/g, "").trim(), model };
}
