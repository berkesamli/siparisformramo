import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import {
  whatsappConfigured,
  patronNumaralari,
  sablonAdi,
  sablonDili,
  sendPdfToPatron,
} from "@/lib/whatsapp-pdf";
import { generateOrderPdf } from "@/lib/order-pdf";

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
    alicilar: patronNumaralari().map(maskele),
    sablon: sablonAdi() || null,
    dil: sablonDili(),
  });
}

// Örnek bir fiş PDF'i üretip patron numaralarına gönderir — kurulumu sınamak için.
export async function POST() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const simdi = new Date();
  const pdf = await generateOrderPdf({
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
  });
  const sonuc = await sendPdfToPatron(pdf, {
    tur: "toptan",
    orderId: "TEST",
    musteri: "Deneme Müşteri",
    tutar: 1421,
    calisan: user.name,
  });
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
