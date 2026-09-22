import { NextRequest, NextResponse } from "next/server";
import { isMesajci } from "@/data/users";
import { konusma, konusmaGuncelle, mesajlar, okunduIsaretle } from "@/lib/mesaj/db";
import { yanitGonder } from "@/lib/mesaj/gonder";
import { hataYaniti, mesajKullanici } from "@/lib/mesaj/yetki";
import { pencereAcik, type KonusmaDurum } from "@/lib/mesaj/tur";
import { getCustomer, customerTitle, musteriBolgesi, bolgeler } from "@/lib/customers";
import { getRetailCustomer } from "@/lib/retail-customers";
import { cariOzet, mikroConfigured } from "@/lib/mikro";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

interface MusteriOzeti {
  id: string; tur: "toptan" | "perakende"; ad: string; telefon: string; eposta: string; sehir?: string; bolge?: string;
  iskontoPct?: number; mikroBakiye?: number | null; mikroUnvan?: string; href: string;
}

async function musteriOzeti(musteriId: string | null, tur: "toptan" | "perakende" | null): Promise<MusteriOzeti | null> {
  if (!musteriId) return null;
  if (tur === "perakende") {
    const c = await getRetailCustomer(musteriId).catch(() => null);
    return c ? { id: c.id, tur: "perakende", ad: c.name, telefon: c.phone, eposta: c.email, href: "/panel/perakende/musteriler" } : null;
  }
  const c = await getCustomer(musteriId).catch(() => null);
  if (!c) return null;
  const b = musteriBolgesi(c);
  let mikroBakiye: number | null | undefined;
  if (c.mikroCariKod && mikroConfigured()) {
    const r = await cariOzet(c.mikroCariKod).catch(() => null);
    mikroBakiye = r?.ok && r.ozet ? r.ozet.bakiye : null;
  }
  return {
    id: c.id, tur: "toptan", ad: customerTitle(c), telefon: c.phone, eposta: c.email,
    sehir: [c.city, c.district].filter(Boolean).join(" / "), bolge: bolgeler()[b]?.label || b,
    iskontoPct: c.iskontoPct, mikroBakiye, mikroUnvan: c.mikroUnvan, href: `/musteriler/kart?id=${encodeURIComponent(c.id)}`,
  };
}

export async function GET(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  try {
    const k = await konusma(params.id);
    if (!k) return NextResponse.json({ ok: false, error: "Konuşma bulunamadı." }, { status: 404 });
    if (req.nextUrl.searchParams.get("oku") !== "0" && k.okunmamis > 0) { await okunduIsaretle(k.id); k.okunmamis = 0; }
    const [ms, musteri] = await Promise.all([mesajlar(k.id), musteriOzeti(k.musteriId, k.musteriTur)]);
    return NextResponse.json({ ok: true, konusma: k, mesajlar: ms, musteri, pencere: pencereAcik(k) });
  } catch (e) {
    return hataYaniti(e);
  }
}

// Durum / atama / müşteri bağlama / ad
export async function PATCH(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as { durum?: string; atanan?: string | null; musteriId?: string | null; musteriTur?: string | null; ad?: string } | null;
  if (!b) return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  const d: Parameters<typeof konusmaGuncelle>[1] = {};
  if (b.durum !== undefined) {
    if (!["acik", "yanitlandi", "kapali"].includes(String(b.durum))) return NextResponse.json({ ok: false, error: "Geçersiz durum." }, { status: 400 });
    d.durum = b.durum as KonusmaDurum;
  }
  if (b.atanan !== undefined) {
    const a = String(b.atanan || "").trim();
    if (a && !isMesajci(a)) return NextResponse.json({ ok: false, error: "Bu kullanıcıya mesaj yetkisi tanımlı değil." }, { status: 400 });
    d.atanan = a || null;
  }
  if (b.musteriId !== undefined) {
    const id = String(b.musteriId || "").trim();
    d.musteriId = id || null;
    d.musteriTur = id ? (b.musteriTur === "perakende" || id.startsWith("P") ? "perakende" : "toptan") : null;
  }
  if (b.ad !== undefined) d.ad = String(b.ad).trim().slice(0, 120);
  try {
    const k = await konusmaGuncelle(params.id, d);
    if (!k) return NextResponse.json({ ok: false, error: "Konuşma bulunamadı." }, { status: 404 });
    return NextResponse.json({ ok: true, konusma: k, musteri: await musteriOzeti(k.musteriId, k.musteriTur) });
  } catch (e) {
    return hataYaniti(e);
  }
}

// Yanıt gönder
export async function POST(req: NextRequest, { params }: { params: { id: string } }) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as { metin?: string; taslakAi?: boolean } | null;
  if (!b?.metin?.trim()) return NextResponse.json({ ok: false, error: "Mesaj metni boş." }, { status: 400 });
  try {
    const r = await yanitGonder(params.id, b.metin, { username: y.user.username, name: y.user.name }, Boolean(b.taslakAi));
    const k = await konusma(params.id);
    return NextResponse.json({ ok: true, mesaj: r.mesaj, yontem: r.yontem, sablon: r.sablon, konusma: k });
  } catch (e) {
    console.error("Mesaj gönderilemedi:", e);
    return hataYaniti(e, 502);
  }
}
