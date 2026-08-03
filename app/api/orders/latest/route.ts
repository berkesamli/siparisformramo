import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { blobConfigured } from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Yeni sipariş bildirimi için hafif durum ucu: yalnızca iki sayaç dosyası
// okunur (toptan + perakende). Sipariş listesini çekmekten çok daha ucuzdur —
// panel açıkken belirli aralıklarla çağrılır, Blob işlem limitini yormaz.

async function readSeq(path: string): Promise<number> {
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path, { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return 0;
    const d = JSON.parse(await new Response(r.stream).text());
    return Number(d.seq) || 0;
  } catch {
    return 0;
  }
}

export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  if (!blobConfigured()) {
    return NextResponse.json({ ok: true, toptan: 0, perakende: 0 });
  }
  const year = new Date().toLocaleDateString("en-CA", {
    timeZone: "Europe/Istanbul",
  }).slice(0, 4);
  const [toptan, perakende] = await Promise.all([
    readSeq(`orders/counter-${year}.json`),
    readSeq(`retail/counter-${year}.json`),
  ]);
  return NextResponse.json({ ok: true, toptan, perakende });
}
