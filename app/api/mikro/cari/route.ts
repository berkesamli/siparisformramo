import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getCustomer, saveCustomer } from "@/lib/customers";
import { mikroConfigured, cariAra, cariOzet } from "@/lib/mikro";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

// Çalışanların kullandığı Mikro cari uçları (yalnızca okuma + müşteri kartına
// eşleştirme kaydı). Sorgular lib/mikro.ts içinde sabittir.
//   GET ?q=ad          → Mikro'da cari araması
//   GET ?musteri=C123  → müşterinin bağlı Mikro carisi ve canlı bakiyesi
//   POST { musteri, cariKod, unvan } → eşleştir · cariKod boş → bağlantıyı kaldır

const MUSTERI_ID = /^[A-Za-z0-9]{2,40}$/;
const CARI_KOD = /^[A-Za-z0-9._\-/ ]{1,40}$/;

export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz" }, { status: 401 });
  }
  if (!mikroConfigured()) {
    return NextResponse.json({ ok: false, kurulu: false, error: "Mikro bağlantısı ayarlanmamış." });
  }
  const sp = req.nextUrl.searchParams;
  const q = sp.get("q");
  if (q !== null) {
    const r = await cariAra(q);
    return NextResponse.json({ ...r, kurulu: true });
  }
  const musteri = sp.get("musteri") || "";
  if (!MUSTERI_ID.test(musteri)) {
    return NextResponse.json({ ok: false, error: "Geçersiz müşteri" }, { status: 400 });
  }
  const c = await getCustomer(musteri);
  if (!c) return NextResponse.json({ ok: false, error: "Müşteri bulunamadı" }, { status: 404 });
  // Arama kutusu için öneri: firma adı, yoksa kişi adı
  const oneri = (c.company || `${c.firstName} ${c.lastName}`).trim();
  if (!c.mikroCariKod) {
    return NextResponse.json({ ok: true, kurulu: true, bagli: false, oneri });
  }
  const r = await cariOzet(c.mikroCariKod);
  return NextResponse.json({
    ok: r.ok,
    kurulu: true,
    bagli: true,
    cariKod: c.mikroCariKod,
    unvan: c.mikroUnvan || r.ozet?.unvan || "",
    ozet: r.ozet,
    hata: r.hata,
    oneri,
  });
}

export async function POST(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz" }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as { musteri?: string; cariKod?: string; unvan?: string } | null;
  const musteri = String(body?.musteri || "");
  if (!MUSTERI_ID.test(musteri)) {
    return NextResponse.json({ ok: false, error: "Geçersiz müşteri" }, { status: 400 });
  }
  const cariKod = String(body?.cariKod || "").trim().slice(0, 40);
  const unvan = String(body?.unvan || "").trim().slice(0, 120);
  if (cariKod && !CARI_KOD.test(cariKod)) {
    return NextResponse.json({ ok: false, error: "Geçersiz cari kodu" }, { status: 400 });
  }
  const c = await getCustomer(musteri);
  if (!c) return NextResponse.json({ ok: false, error: "Müşteri bulunamadı" }, { status: 404 });
  const guncel = {
    ...c,
    mikroCariKod: cariKod || undefined,
    mikroUnvan: cariKod ? unvan || undefined : undefined,
    updatedAt: new Date().toISOString(),
  };
  const kaydedildi = await saveCustomer(guncel);
  if (!kaydedildi) {
    return NextResponse.json({ ok: false, error: "Kalıcı depolama yapılandırılmadığı için eşleştirme kaydedilemedi." }, { status: 503 });
  }
  return NextResponse.json({ ok: true, customer: guncel });
}
