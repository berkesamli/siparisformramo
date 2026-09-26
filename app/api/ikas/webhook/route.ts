import { NextRequest, NextResponse } from "next/server";
import { createHmac, timingSafeEqual } from "node:crypto";
import { dbConfigured } from "@/lib/uretim/db";
import { ikasWebhookKey } from "@/lib/ikas/client";
import { ikasWebhookIsle } from "@/lib/ikas/senk";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// ikas webhook alıcısı (store/order/created, store/order/updated).
// Gövde: { id, scope, merchantId, authorizedAppId, createdAt, data: "<sipariş JSON metni>", signature }.
// İmza = hex(HMAC-SHA256(client secret, data metni)). İmza doğruysa ya da adresteki
// IKAS_WEBHOOK_KEY (?k=… / x-webhook-key) tutuyorsa kabul edilir; sipariş yine de
// API'den yeniden okunur (yükteki veri yalnızca tetikleyici).
// ikas 200 dışı yanıtta 3 kez dener ve hata oranı yüksek adresi engeller; bu yüzden
// işleme hatası da 200 ile döner (senkron cron/açılışta yakalar).
function esit(a: string, b: string): boolean {
  const x = Buffer.from(a), y = Buffer.from(b);
  return x.length === y.length && x.length > 0 && timingSafeEqual(x, y);
}

function anahtarDogru(req: NextRequest): boolean {
  const beklenen = ikasWebhookKey();
  if (!beklenen) return false;
  return esit(req.nextUrl.searchParams.get("k") || req.headers.get("x-webhook-key") || "", beklenen);
}

function imzaDogru(govde: any): boolean {
  const gizli = (process.env.IKAS_CLIENT_SECRET || "").trim();
  if (!gizli || typeof govde?.data !== "string" || typeof govde?.signature !== "string") return false;
  const hesap = createHmac("sha256", gizli).update(govde.data, "utf8").digest("hex");
  return esit(hesap, String(govde.signature).toLowerCase());
}

export async function POST(req: NextRequest) {
  const govde = (await req.json().catch(() => null)) as any;
  if (!govde) return NextResponse.json({ ok: false, error: "Geçersiz gövde." }, { status: 400 });
  if (!imzaDogru(govde) && !anahtarDogru(req)) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  if (!dbConfigured()) return NextResponse.json({ ok: true, atlandi: "DATABASE_URL yok." });
  let veri: any = govde.data ?? govde;
  if (typeof veri === "string") { try { veri = JSON.parse(veri); } catch { veri = null; } }
  const scope = String(govde.scope || req.headers.get("x-ikas-event") || "");
  if (!veri || typeof veri !== "object" || !veri.id) return NextResponse.json({ ok: true, atlandi: "Sipariş kimliği yok.", scope });
  try {
    const r = await ikasWebhookIsle(veri);
    return NextResponse.json({ ok: true, scope, sonuc: r });
  } catch (e) {
    console.error("ikas webhook işlenemedi:", e);
    return NextResponse.json({ ok: false, scope, error: (e as Error)?.message || String(e) });
  }
}

// Bazı sistemler webhook adresini GET ile doğrular
export async function GET(req: NextRequest) {
  return NextResponse.json({ ok: anahtarDogru(req), servis: "olga-uretim-takvimi ikas webhook" });
}
