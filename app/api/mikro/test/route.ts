import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import { mikroAyar, mikroConfigured, baglantiTesti, cariOzet, sifreBicimiAdi, mikroAyarUyarilari } from "@/lib/mikro";

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
    sifreUzunluk: a.sifre.length,
    uyarilar: mikroAyarUyarilari(),
  });
}

// { islem: "baglanti" } → ilk 5 cari kart · { islem: "cari", cariKod } → bakiye özeti
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as {
    islem?: string; cariKod?: string;
    // "farklı bilgilerle dene": yalnızca bu istekte kullanılır, hiçbir yere kaydedilmez
    dene?: { kullanici?: string; sifre?: string; firma?: string; yil?: string };
  } | null;
  if (body?.islem === "cari") {
    const r = await cariOzet(String(body.cariKod || "").slice(0, 40));
    return NextResponse.json(r);
  }
  const d = body?.dene;
  const override = d && (d.kullanici || d.sifre || d.firma || d.yil)
    ? { kullanici: String(d.kullanici || "").slice(0, 40), sifre: String(d.sifre || "").slice(0, 80), firma: String(d.firma || "").slice(0, 40), yil: String(d.yil || "").slice(0, 4) }
    : undefined;
  const r = await baglantiTesti(override);
  // Tutan varyant başarılı yanıttan gelir; tutmadıysa bunu açıkça söyle (eskiden ilk biçim yazılıyordu, yanıltıyordu)
  const sifreBicimi = r.bicim
    ? (override ? `deneme bilgileri · ${r.bicim}` : r.bicim)
    : (override ? "deneme bilgileri · hiçbir kullanıcı/şifre biçimi tutmadı" : `hiçbir kullanıcı/şifre biçimi tutmadı (ilk denenen: ${sifreBicimiAdi()})`);
  return NextResponse.json({ ...r, sifreBicimi });
}
