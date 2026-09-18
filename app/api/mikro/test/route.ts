import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import { mikroAyar, mikroConfigured, baglantiTesti, cariOzet, sifreBicimiAdi } from "@/lib/mikro";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

// Mikro bağlantı durumu ve sabit deneme sorguları — yalnızca sahipler.
export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const a = mikroAyar();
  return NextResponse.json({
    ok: true,
    kurulu: mikroConfigured(),
    url: a.url || null,
    firma: a.firma || null,
    kullanici: a.kullanici || null,
    yil: a.yil,
    apiKey: Boolean(a.apiKey),
    sifre: Boolean(a.sifre),
  });
}

// { islem: "baglanti" } → ilk 5 cari kart · { islem: "cari", cariKod } → bakiye özeti
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as { islem?: string; cariKod?: string } | null;
  if (body?.islem === "cari") {
    const r = await cariOzet(String(body.cariKod || "").slice(0, 40));
    return NextResponse.json(r);
  }
  const r = await baglantiTesti();
  return NextResponse.json({ ...r, sifreBicimi: sifreBicimiAdi() });
}
