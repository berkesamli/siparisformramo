import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { toptanIndeks } from "@/lib/dashboard";
import { blobConfigured, istanbulDateKey, lastNDateKeys, type OrderIndexEntry, type OrderStatus } from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 20;

// Ana sayfa listeleri: açık toptan siparişler (en yeni önce) ve son 7 günde
// merkez kontrolü yapılmamış siparişler. Aylık indeks üzerinden çalışır
// (gösterge paneliyle aynı önbellek); sipariş dosyaları tek tek okunmaz.

const ACIK = new Set<OrderStatus>(["olusturuldu", "hazirlaniyor", "yarim"]);

export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  try {
    const idx = await toptanIndeks();
    const set7 = new Set(lastNDateKeys(7));
    const yeniOnce = (a: OrderIndexEntry, b: OrderIndexEntry) => b.createdAt.localeCompare(a.createdAt);
    const acikHepsi = idx.filter((o) => ACIK.has(o.status)).sort(yeniOnce);
    const kontrolsuz = idx
      .filter((o) => set7.has(o.dateKey) && o.status !== "iptal" && o.kontrol === false)
      .sort(yeniOnce)
      .slice(0, 12);
    return NextResponse.json({
      ok: true,
      blob: blobConfigured(),
      today: istanbulDateKey(),
      acik: acikHepsi.slice(0, 15),
      toplamAcik: acikHepsi.length,
      kontrolsuz,
    });
  } catch (err) {
    console.error("Açık sipariş listesi alınamadı:", err);
    return NextResponse.json({ ok: false, error: "Veriler alınamadı." }, { status: 500 });
  }
}
