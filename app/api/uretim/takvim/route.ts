import { NextRequest, NextResponse } from "next/server";
import { uretimKullanici, hataYaniti } from "@/lib/uretim/yetki";
import { isler, notlar, teklifler } from "@/lib/uretim/db";
import { perakendeSenkGerekirse } from "@/lib/uretim/perakende-senk";
import { istanbulDateKey } from "@/lib/orders";
import { TARIH_RE, type IsKaynak, type Sube, type TakvimYaniti } from "@/lib/uretim/tur";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Takvim verisi: aralıktaki planlı işler + (plansiz=1) planlanmamış açık işler +
// gün notları + takip tarihi aralıkta olan açık teklifler. Açılışta mağaza
// (perakende) ve online (ikas) siparişleri gerekirse senkronlanır (senk=0 ile kapatılır).
export async function GET(req: NextRequest) {
  const y = await uretimKullanici();
  if ("hata" in y) return y.hata;
  const q = req.nextUrl.searchParams;
  const from = q.get("from") || "";
  const to = q.get("to") || "";
  if (!TARIH_RE.test(from) || !TARIH_RE.test(to) || from > to) {
    return NextResponse.json({ ok: false, error: "Geçersiz tarih aralığı." }, { status: 400 });
  }
  const kaynak = (q.get("kaynak") || "") as IsKaynak | "";
  const sube = (q.get("sube") || "hepsi") as Sube | "hepsi";
  const plansiz = q.get("plansiz") !== "0";
  const senk: TakvimYaniti["senk"] = {};
  if (q.get("senk") !== "0") {
    try {
      const r = await perakendeSenkGerekirse(y.user.name);
      if (r) senk.perakende = { eklenen: r.eklenen, guncellenen: r.guncellenen };
    } catch (e) {
      console.error("Perakende senkronu:", e);
    }
    try {
      const { ikasSenkGerekirse } = await import("@/lib/ikas/senk");
      const r = await ikasSenkGerekirse(y.user.name);
      if (r) senk.online = { eklenen: r.eklenen, guncellenen: r.guncellenen, hata: r.hata };
    } catch (e) {
      console.error("ikas senkronu:", e);
    }
  }
  try {
    const [isListesi, notListesi, teklifListesi] = await Promise.all([
      isler({ from, to, plansiz, kaynak: kaynak || undefined, sube, q: q.get("q") || undefined }),
      notlar(from, to),
      teklifler({ durum: "acik", takipFrom: from, takipTo: to }),
    ]);
    const yanit: TakvimYaniti = { ok: true, from, to, bugun: istanbulDateKey(), isler: isListesi, notlar: notListesi, teklifler: teklifListesi, senk };
    return NextResponse.json(yanit);
  } catch (e) {
    console.error("Takvim okunamadı:", e);
    return hataYaniti(e);
  }
}
