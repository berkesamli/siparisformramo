import { NextResponse } from "next/server";
import { createHmac, timingSafeEqual } from "node:crypto";
import { dbConfigured } from "@/lib/mesaj/db";
import { gelenInstagram, type IgEntry } from "@/lib/mesaj/instagram";
import { gelenWhatsapp, whatsappDurumlar, type WaValue, type WaDurum } from "@/lib/mesaj/whatsapp";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Meta webhook'u — Instagram mesajları (object: "instagram"). Aynı Meta
// uygulamasındaki WhatsApp olayları da buraya yönlendirilirse işlenir.
//   GET  — doğrulama (hub.verify_token = INSTAGRAM_VERIFY_TOKEN ya da WHATSAPP_VERIFY_TOKEN)
//   POST — olaylar; META_APP_SECRET / WHATSAPP_APP_SECRET tanımlıysa imza doğrulanır.

const verifyToken = () => (process.env.INSTAGRAM_VERIFY_TOKEN || process.env.WHATSAPP_VERIFY_TOKEN || "").trim();
const appSecret = () => (process.env.META_APP_SECRET || process.env.WHATSAPP_APP_SECRET || "").trim();

export async function GET(req: Request) {
  const u = new URL(req.url);
  const beklenen = verifyToken();
  if (u.searchParams.get("hub.mode") === "subscribe" && beklenen && u.searchParams.get("hub.verify_token") === beklenen) {
    return new NextResponse(u.searchParams.get("hub.challenge") || "", { status: 200, headers: { "Content-Type": "text/plain" } });
  }
  return NextResponse.json({ ok: false, error: "Doğrulama başarısız." }, { status: 403 });
}

function imzaGecerli(raw: string, header: string | null): boolean {
  const secret = appSecret();
  if (!secret) return true;
  if (!header || !header.startsWith("sha256=")) return false;
  const beklenen = createHmac("sha256", secret).update(raw, "utf8").digest("hex");
  const gelen = header.slice(7);
  return gelen.length === beklenen.length && timingSafeEqual(Buffer.from(gelen, "hex"), Buffer.from(beklenen, "hex"));
}

interface Govde {
  object?: string;
  entry?: (IgEntry & { changes?: { field?: string; value?: WaValue & { statuses?: WaDurum[] } }[] })[];
}

export async function POST(req: Request) {
  const raw = await req.text();
  if (!imzaGecerli(raw, req.headers.get("x-hub-signature-256"))) {
    return NextResponse.json({ ok: false, error: "İmza geçersiz." }, { status: 401 });
  }
  let body: Govde | null = null;
  try { body = JSON.parse(raw); } catch { return NextResponse.json({ ok: false }, { status: 400 }); }
  if (!dbConfigured()) {
    console.warn("Mesaj webhook'u geldi ama DATABASE_URL yok; olay atlandı.");
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
  } catch (e) {
    // Meta 200 dışında yanıt alırsa tekrar dener; işlenen kısmı kaybetmemek için hata loglanır, 200 döner.
    console.error("Mesaj webhook işlenemedi:", e);
  }
  return NextResponse.json({ ok: true, yeni, durum });
}
