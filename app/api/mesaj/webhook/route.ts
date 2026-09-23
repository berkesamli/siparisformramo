import { NextResponse } from "next/server";
import { createHmac, timingSafeEqual } from "node:crypto";
import { dbConfigured } from "@/lib/mesaj/db";
import { gelenInstagram, type IgEntry } from "@/lib/mesaj/instagram";
import { gelenWhatsapp, whatsappDurumlar, type WaValue, type WaDurum } from "@/lib/mesaj/whatsapp";
import { izKaydet } from "@/lib/mesaj/webhook-iz";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Meta webhook'u — Instagram mesajları (object: "instagram"). Aynı Meta
// uygulamasındaki WhatsApp olayları da buraya yönlendirilirse işlenir.
//   GET  — doğrulama (hub.verify_token = INSTAGRAM_VERIFY_TOKEN ya da WHATSAPP_VERIFY_TOKEN)
//   POST — olaylar; imza ZORUNLU: INSTAGRAM_APP_SECRET (Instagram login yolu), META_APP_SECRET ya da
//          WHATSAPP_APP_SECRET ile doğrulanır (tanımlı olanların hepsi denenir); hiçbiri yoksa 401.

const verifyToken = () => (process.env.INSTAGRAM_VERIFY_TOKEN || process.env.WHATSAPP_VERIFY_TOKEN || "").trim();
const secretler = () => [...new Set([process.env.INSTAGRAM_APP_SECRET, process.env.META_APP_SECRET, process.env.WHATSAPP_APP_SECRET].map((s) => (s || "").trim()).filter(Boolean))];
const appSecret = () => secretler()[0] || "";

export async function GET(req: Request) {
  const u = new URL(req.url);
  const beklenen = verifyToken();
  if (u.searchParams.get("hub.mode") === "subscribe" && beklenen && u.searchParams.get("hub.verify_token") === beklenen) {
    return new NextResponse(u.searchParams.get("hub.challenge") || "", { status: 200, headers: { "Content-Type": "text/plain" } });
  }
  return NextResponse.json({ ok: false, error: "Doğrulama başarısız." }, { status: 403 });
}

function imzaGecerli(raw: string, header: string | null): boolean {
  const liste = secretler();
  if (!liste.length) return false; // gizli anahtar girilmeden POST kabul edilmez (fail-closed)
  if (!header || !/^sha256=[0-9a-f]{64}$/i.test(header)) return false;
  const gelen = Buffer.from(header.slice(7), "hex");
  // Aynı uygulamada Instagram login (Instagram app secret) ve WhatsApp (Facebook app secret) farklı anahtarla imzalar.
  return liste.some((s) => timingSafeEqual(gelen, createHmac("sha256", s).update(raw, "utf8").digest()));
}

interface Govde {
  object?: string;
  entry?: (IgEntry & { changes?: { field?: string; value?: WaValue & { statuses?: WaDurum[] } }[] })[];
}

/** Instagram olaylarının türlere göre dökümü: "3 olay (mesaj 2, okundu 1)"; standby varsa uyarır. */
function olayDokumu(body: Govde | null): { n: number; ozet: string } {
  const say: Record<string, number> = {};
  let standby = 0, n = 0;
  for (const e of body?.entry || []) {
    standby += e.standby?.length || 0;
    for (const ev of e.messaging || []) {
      n++;
      const tur = ev.message
        ? (ev.message.is_deleted ? "silindi" : ev.message.is_echo ? "bizim gönderdiğimiz (echo)" : ev.message.is_unsupported ? "desteklenmeyen" : "mesaj")
        : ev.read ? "okundu" : ev.reaction ? "tepki" : ev.postback ? "postback" : "diğer";
      say[tur] = (say[tur] || 0) + 1;
    }
  }
  const parcalar = Object.entries(say).map(([t, c]) => `${t} ${c}`).join(", ");
  let ozet = `object: ${body?.object || "?"} · ${n} olay${parcalar ? ` (${parcalar})` : ""}`;
  if (standby) ozet += ` · standby ${standby}: Meta bu mesajları başka bir uygulamaya "birincil alıcı" olarak veriyor; Meta uygulaması → Messenger ayarları → Handover'da bu uygulamayı birincil yapın ya da diğer uygulamanın sayfa aboneliğini kaldırın`;
  return { n, ozet };
}

export async function POST(req: Request) {
  const raw = await req.text();
  if (!imzaGecerli(raw, req.headers.get("x-hub-signature-256"))) {
    await izKaydet("instagram", "imza-red", appSecret()
      ? "Meta'dan olay geldi ama imza doğrulanamadı: Instagram login yolunda Instagram uygulama gizli anahtarı (INSTAGRAM_APP_SECRET; Instagram API → Instagram app secret) gerekir; Facebook sayfası yolunda META_APP_SECRET / WHATSAPP_APP_SECRET."
      : "Meta'dan olay geldi ama hiçbir gizli anahtar (INSTAGRAM_APP_SECRET / META_APP_SECRET / WHATSAPP_APP_SECRET) tanımlı olmadığı için reddedildi.");
    return NextResponse.json({ ok: false, error: "İmza geçersiz." }, { status: 401 });
  }
  let body: Govde | null = null;
  try { body = JSON.parse(raw); } catch { return NextResponse.json({ ok: false }, { status: 400 }); }
  const kanal = body?.object === "instagram" ? "instagram" : "whatsapp";
  const dokum = olayDokumu(body);
  await izKaydet(kanal, dokum.n ? "mesaj" : "diger", dokum.ozet);
  if (!dbConfigured()) {
    console.warn("Mesaj webhook'u geldi ama DATABASE_URL yok; olay atlandı.");
    await izKaydet(kanal, "islem-hata", "Olay işlenemedi: DATABASE_URL tanımlı değil (gelen kutusu kurulu değil).");
    return NextResponse.json({ ok: true, atlandi: true });
  }
  let yeni = 0, durum = 0;
  try {
    for (const entry of body?.entry || []) {
      if (body?.object === "instagram") {
        yeni += await gelenInstagram(entry);
      } else if (body?.object === "whatsapp_business_account") {
        for (const ch of entry.changes || []) {
          const v = ch.value;
          if (!v) continue;
          if (v.messages?.length) yeni += await gelenWhatsapp(v);
          if (v.statuses?.length) durum += await whatsappDurumlar(v.statuses);
        }
      }
    }
    await izKaydet(kanal, "islem", `${dokum.n} olay → ${yeni} yeni mesaj${durum ? `, ${durum} durum güncellemesi` : ""}${dokum.n && !yeni && !durum ? " (yeni mesaj yok: okundu/tepki olayı ya da daha önce kaydedilmiş mesaj)" : ""}`);
  } catch (e) {
    // Meta 200 dışında yanıt alırsa tekrar dener; işlenen kısmı kaybetmemek için hata loglanır, 200 döner.
    console.error("Mesaj webhook işlenemedi:", e);
    await izKaydet(kanal, "islem-hata", `Olay işlenirken hata: ${(e as Error)?.message || String(e)}`);
  }
  return NextResponse.json({ ok: true, yeni, durum });
}
