import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import { mikroConfigured } from "@/lib/mikro";
import { mikroStokCek, mikroStokCekVeKaydet } from "@/lib/mikro-stok";
import { bust } from "@/lib/server-cache";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Mikro'dan stok çekme (günlük Excel'in yerine).
//   POST { kaydet?: boolean }  çalışanlar — kaydet=false yalnızca önizleme (sahipler)
//   GET                        Vercel cron (Authorization: Bearer CRON_SECRET) ya da sahip oturumu → çek + kaydet

function ozet(r: Awaited<ReturnType<typeof mikroStokCek>>, kaydedildi: boolean) {
  const items = r.data?.items || [];
  return {
    ok: r.ok,
    error: r.ok ? undefined : r.hata,
    kaydedildi,
    count: items.length,
    ankaraTotal: Math.round(items.reduce((s, i) => s + i.ankaraMt, 0)),
    istanbulTotal: Math.round(items.reduce((s, i) => s + i.istanbulMt, 0)),
    updatedAt: r.data?.updatedAt,
    sourceName: r.data?.sourceName,
    bilgi: r.bilgi,
  };
}

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  if (!mikroConfigured()) {
    return NextResponse.json({ ok: false, error: "Mikro bağlantısı ayarlanmamış (Bildirim Ayarları → Mikro Bağlantısı)." }, { status: 503 });
  }
  const body = (await req.json().catch(() => ({}))) as { kaydet?: boolean };
  const kaydet = body?.kaydet !== false;
  if (!kaydet && !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Önizleme yalnızca sahipler için." }, { status: 403 });
  }
  const r = kaydet ? await mikroStokCekVeKaydet() : { ...(await mikroStokCek()), kaydedildi: false };
  if (r.kaydedildi) { bust("dash:stok"); bust("search:stok"); }
  return NextResponse.json(ozet(r, r.kaydedildi), { status: r.ok ? 200 : 502 });
}

export async function GET(req: Request) {
  const sir = process.env.CRON_SECRET;
  const auth = req.headers.get("authorization") || "";
  let yetkili = Boolean(sir && auth === `Bearer ${sir}`);
  if (!yetkili) {
    const user = await getSessionUser();
    yetkili = !!user && user.role === "staff" && isOwner(user.username);
  }
  if (!yetkili) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  if (!mikroConfigured()) return NextResponse.json({ ok: false, error: "Mikro bağlantısı ayarlanmamış." }, { status: 503 });
  const r = await mikroStokCekVeKaydet();
  if (r.kaydedildi) { bust("dash:stok"); bust("search:stok"); }
  if (!r.ok) console.error("Mikro stok çekimi başarısız:", r.hata);
  return NextResponse.json(ozet(r, r.kaydedildi), { status: r.ok ? 200 : 502 });
}
