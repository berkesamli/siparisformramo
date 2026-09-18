import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import {
  whatsappConfigured,
  patronAlicilar,
  sablonAdi,
  sablonDili,
  sendPdfToPatron,
  musteriSablonAdlari,
  sendPdfToCustomer,
} from "@/lib/whatsapp-pdf";
import { generateOrderPdf, musteriKopyasi } from "@/lib/order-pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

const maskele = (n: string) => n.replace(/^(\d{2})(\d{3})\d+(\d{2})$/, "$1 $2 *** ** $3");

// Bildirim ayarlarının durumu — yalnızca sahipler.
export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  return NextResponse.json({
    ok: true,
    api: whatsappConfigured(),
    alicilar: patronAlicilar().map((a) => `${a.ad} · ${maskele(a.to)}`),
    sablon: sablonAdi() || null,
    dil: sablonDili(),
    musteriSablon: musteriSablonAdlari().join(", ") || null,
    webhook: Boolean((process.env.WHATSAPP_VERIFY_TOKEN || "").trim()),
    imza: Boolean((process.env.WHATSAPP_APP_SECRET || "").trim()),
  });
}

// Örnek bir fiş PDF'i üretip gönderir — kurulumu sınamak için.
//   gövde yok / {}                → patron numaralarına (PATRON_WHATSAPP)
//   { patronTelefon: "05…" }      → o numaraya patron şablonuyla
//   { musteriTelefon: "05…" }     → o numaraya müşteri şablonuyla
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as { musteriTelefon?: string; patronTelefon?: string } | null;
  const musteriTelefon = String(body?.musteriTelefon || "").trim();
  const patronTelefon = String(body?.patronTelefon || "").trim();
  const simdi = new Date();
  const ornek = {
    orderId: "TEST-" + simdi.toISOString().slice(11, 16).replace(":", ""),
    dateStr: simdi.toLocaleString("tr-TR", { timeZone: "Europe/Istanbul" }),
    status: "Test",
    employee: user.name,
    customer: "Deneme Müşteri",
    note: "Bu bir test fişidir — WhatsApp bildirim kurulumu doğrulanıyor.",
    discountPct: 0,
    vatApplied: false,
    lines: [{ name: "GC065-1473 (Çerçeve Profili)", unitText: "29 mt (10 boy)", unitPriceTL: 49, lineTotal: 1421 }],
    gross: 1421,
    discount: 0,
    vatAmount: 0,
    net: 1421,
  };
  const pdf = await generateOrderPdf(ornek);
  if (musteriTelefon) {
    const musteriPdf = await generateOrderPdf(musteriKopyasi(ornek));
    const w = await sendPdfToCustomer(musteriPdf, { telefon: musteriTelefon, musteri: "Deneme Müşteri", orderId: "TEST" });
    return NextResponse.json({
      ok: w.ok,
      sonuc: {
        ok: w.ok,
        gonderilen: w.ok && w.to ? [maskele(w.to)] : [],
        hatalar: w.hata ? [w.hata] : [],
        notlar: w.ok
          ? [...(w.not ? [w.not] : []), "Meta mesajı kabul etti. Numarada WhatsApp yoksa teslim hatası webhook'la gelir ve sistem SMS'e düşer."]
          : [],
        yontem: w.yontem,
        sablon: w.sablon,
      },
    });
  }
  const sonuc = await sendPdfToPatron(
    pdf,
    { tur: "toptan", orderId: "TEST", musteri: "Deneme Müşteri", tutar: 1421, calisan: user.name },
    patronTelefon ? [patronTelefon] : undefined
  );
  // Ekranda numaralar maskelenir (gidenler, notlar ve hatalar "numara: ..." ile başlar).
  const maskeleSatir = (s: string) => s.replace(/^(\d{11,15})(?=:)/, maskele);
  return NextResponse.json({
    ok: sonuc.ok,
    sonuc: {
      ...sonuc,
      gonderilen: sonuc.gonderilen.map(maskele),
      notlar: sonuc.notlar.map(maskeleSatir),
      hatalar: sonuc.hatalar.map(maskeleSatir),
    },
  });
}
