import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { isBul, isGuncelle, isSirala } from "@/lib/uretim/db";
import { TARIH_RE } from "@/lib/uretim/tur";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Sürükle-bırak: { id, tarih, ids } — `id` işi `tarih` gününe taşır (olay yazılır), sonra o
// günün sırası `ids` dizisine göre yazılır. tarih null → planlanmamış kuyruğa.
export async function POST(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as { id?: string; tarih?: string | null; ids?: string[] } | null;
  if (!b || !b.id) return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  const tarih = b.tarih ? String(b.tarih).slice(0, 10) : null;
  if (tarih && !TARIH_RE.test(tarih)) return NextResponse.json({ ok: false, error: "Geçersiz tarih." }, { status: 400 });
  const ids = Array.isArray(b.ids) ? b.ids.filter((x) => typeof x === "string").slice(0, 300) : [];
  try {
    const eski = await isBul(b.id);
    if (!eski) return NextResponse.json({ ok: false, error: "İş bulunamadı." }, { status: 404 });
    if (eski.planTarih !== tarih) await isGuncelle(b.id, { planTarih: tarih }, y.user.name);
    if (tarih && ids.length) await isSirala(tarih, ids.includes(b.id) ? ids : [...ids, b.id], y.user.name);
    return NextResponse.json({ ok: true, is: await isBul(b.id) });
  } catch (e) {
    console.error("Sıralama yazılamadı:", e);
    return hataYaniti(e);
  }
}
