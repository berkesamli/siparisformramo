import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner, mesajUsernames } from "@/data/users";
import { dbConfigured, dbSaglik, konusmalar } from "@/lib/mesaj/db";
import { gmailConfigured, gmailHesaplar, gmailSenk, gmailTest } from "@/lib/mesaj/gmail";
import { igJetonTazele, igSayfaAbonelik, igThreadSahibi, instagramConfigured, instagramDurum, instagramSenk } from "@/lib/mesaj/instagram";
import { igUygulamaWebhookOnar, uygulamaWebhookDurumu } from "@/lib/mesaj/meta-uygulama";
import { SABLON_GOVDE, SABLON_VARSAYILAN, sablonListesi, sablonOlustur, serbestSablonAdlari, wabaAbonelik, whatsappConfigured, whatsappDurum, type SablonBilgi } from "@/lib/mesaj/whatsapp";
import { taslakHazir } from "@/lib/mesaj/taslak";
import { izOku } from "@/lib/mesaj/webhook-iz";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Mesajlar kurulum durumu ve canlı bağlantı testleri — yalnızca sahipler (Ayarlar sayfası).
async function yetki() {
  const u = await getSessionUser();
  return u && u.role === "staff" && isOwner(u.username) ? u : null;
}

export async function GET(req: NextRequest) {
  if (!(await yetki())) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  const canli = req.nextUrl.searchParams.get("test") === "1";
  const [db, gmail, wa, ig, waIz, igIz, abonelik] = await Promise.all([
    canli ? dbSaglik() : Promise.resolve(null),
    canli && gmailConfigured() ? gmailTest() : Promise.resolve(null),
    canli && whatsappConfigured() ? whatsappDurum() : Promise.resolve(null),
    canli && instagramConfigured() ? instagramDurum() : Promise.resolve(null),
    izOku("whatsapp"),
    izOku("instagram"),
    canli && whatsappConfigured() && process.env.WHATSAPP_WABA_ID ? wabaAbonelik(false) : Promise.resolve(null),
  ]);
  const origin = req.nextUrl.origin;
  const igCallback = `${origin}/api/mesaj/webhook`;
  // Handover teşhisi: son Instagram konuşmalarının kontrolü hangi uygulamada? (Facebook sayfası yolu)
  const igThread = canli && instagramConfigured() && (process.env.INSTAGRAM_PAGE_ID || "").trim() && dbConfigured()
    ? await (async () => {
        try {
          const son = await konusmalar({ kanal: "instagram", limit: 3 }, "");
          return await Promise.all(son.map(async (k) => ({ ad: k.ad || k.disKimlik, disKimlik: k.disKimlik, standby: Boolean(k.meta?.standby), sahip: await igThreadSahibi(k.disKimlik) })));
        } catch { return null; }
      })()
    : null;
  // Sınamada Conversations API'den de çekilir: webhook gelmiyorsa mesajlar yine düşer, API'de ne göründüğü kartta yazar.
  const igSenk = canli && instagramConfigured() && (process.env.INSTAGRAM_PAGE_ID || "").trim() && dbConfigured() ? await instagramSenk({ zorla: true }) : null;
  const [igAbonelik, igUygulama, sablonlar] = await Promise.all([
    canli && instagramConfigured() && (process.env.INSTAGRAM_PAGE_ID || "").trim() ? igSayfaAbonelik(false) : Promise.resolve(null),
    canli && instagramConfigured() ? uygulamaWebhookDurumu(igCallback) : Promise.resolve(null),
    canli && whatsappConfigured() ? sablonListesi(true) : Promise.resolve(null),
  ]);
  // Bizim başlattığımız mesajların şablonları: env'dekiler + varsayılan; her birinin Meta'daki durumu.
  const sablonAdaylari = [...new Set([...serbestSablonAdlari(), SABLON_VARSAYILAN])];
  const sablon = sablonlar ? {
    varsayilan: SABLON_VARSAYILAN, govde: SABLON_GOVDE, hata: sablonlar.ok ? undefined : sablonlar.hata,
    adaylar: sablonAdaylari.map((ad): SablonBilgi | { ad: string; durum: "yok" } => sablonlar.liste.find((t) => t.ad === ad && t.dil === (process.env.WHATSAPP_TEMPLATE_DIL || "tr").trim()) || sablonlar.liste.find((t) => t.ad === ad) || { ad, durum: "yok" }),
  } : null;
  // Yayındaki secret'ın kısa parmak izi (tamamı asla dönmez): Meta'daki değerle karşılaştırmak için.
  const secret = (process.env.WHATSAPP_APP_SECRET || process.env.META_APP_SECRET || "").trim();
  const secretIpucu = secret ? { uzunluk: secret.length, bas: secret.slice(0, 2), son: secret.slice(-2), tirnak: /^["']|["']$/.test(secret) } : null;
  return NextResponse.json({
    ok: true,
    db: { kurulu: dbConfigured(), test: db },
    gmail: { kurulu: gmailConfigured(), hesaplar: gmailHesaplar().map((h) => h.adres), test: gmail },
    whatsapp: {
      kurulu: whatsappConfigured(), sablonlar: serbestSablonAdlari(), sablon, test: wa,
      webhook: { url: `${origin}/api/whatsapp/webhook`, verifyToken: Boolean((process.env.WHATSAPP_VERIFY_TOKEN || "").trim()), appSecret: Boolean(secret), secretIpucu, sonOlay: waIz.son, kabul: waIz.kabul, red: waIz.red, islem: waIz.islem, wabaId: Boolean((process.env.WHATSAPP_WABA_ID || "").trim()), abonelik },
    },
    instagram: {
      kurulu: instagramConfigured(), test: ig,
      yol: (process.env.INSTAGRAM_PAGE_ID || "").trim() ? "facebook" : "instagram-login",
      igSecret: Boolean((process.env.INSTAGRAM_APP_SECRET || "").trim()),
      webhook: { url: igCallback, sonOlay: igIz.son, kabul: igIz.kabul, red: igIz.red, islem: igIz.islem, abonelik: igAbonelik, uygulama: igUygulama, thread: igThread, senk: igSenk, sayfa: Boolean((process.env.INSTAGRAM_PAGE_ID || "").trim()) },
    },
    taslak: taslakHazir(),
    gorebilenler: mesajUsernames(),
  });
}

// { islem: "gmail-senk" } → e-postaları şimdi çek
export async function POST(req: NextRequest) {
  if (!(await yetki())) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  const b = (await req.json().catch(() => null)) as { islem?: string } | null;
  if (b?.islem === "ig-jeton") {
    const r = await igJetonTazele(true);
    return NextResponse.json({ ok: r.ok, tazeleme: r, error: r.ok ? undefined : r.hata });
  }
  if (b?.islem === "ig-webhook") {
    const r = await igUygulamaWebhookOnar(`${req.nextUrl.origin}/api/mesaj/webhook`);
    return NextResponse.json({ ok: r.ok, uygulama: r, error: r.ok ? undefined : r.hata });
  }
  if (b?.islem === "ig-abone") {
    const r = await igSayfaAbonelik(true);
    return NextResponse.json({ ok: r.ok, abonelik: r, error: r.ok ? undefined : r.hata });
  }
  if (b?.islem === "wa-sablon") {
    const r = await sablonOlustur();
    return NextResponse.json({ ok: r.ok, sablon: r, error: r.ok ? undefined : r.hata });
  }
  if (b?.islem === "wa-abone") {
    const r = await wabaAbonelik(true);
    return NextResponse.json({ ok: r.ok, abonelik: r, error: r.ok ? undefined : r.hata });
  }
  if (b?.islem === "gmail-senk") {
    if (!dbConfigured()) return NextResponse.json({ ok: false, error: "Önce veri tabanı bağlanmalı (DATABASE_URL)." }, { status: 503 });
    if (!gmailConfigured()) return NextResponse.json({ ok: false, error: "GMAIL_HESAPLAR tanımlı değil." }, { status: 503 });
    const sonuc = await gmailSenk({ zorla: true });
    return NextResponse.json({ ok: true, sonuc });
  }
  return NextResponse.json({ ok: false, error: "Bilinmeyen işlem." }, { status: 400 });
}
