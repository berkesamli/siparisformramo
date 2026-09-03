import { NextRequest, NextResponse } from "next/server";
import { getDealerSession } from "@/lib/auth";
import { getDealer } from "@/lib/dealers";
import { getOrder, verifyTrackingToken } from "@/lib/orders";
import { generateOrderPdf } from "@/lib/pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const TR_ASCII: Record<string, string> = {
  ç: "c", Ç: "C", ğ: "g", Ğ: "G", ı: "i", İ: "I", ö: "o", Ö: "O", ş: "s", Ş: "S", ü: "u", Ü: "U",
};
function asciiFileName(s: string): string {
  return s
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_ASCII[c] || c)
    .replace(/[^A-Za-z0-9 _-]/g, "")
    .trim()
    .replace(/\s+/g, "_");
}

// Üretim PDF'i: bayi kendi oturumuyla, müşteri ise takip linkindeki imzalı
// anahtarla (k) indirir.
export async function GET(req: NextRequest) {
  const q = req.nextUrl.searchParams;
  const d = q.get("d") || "";
  const id = q.get("id") || "";
  const k = q.get("k") || "";
  let slug = q.get("b") || "";

  const sess = await getDealerSession();
  if (sess) slug = sess.dealer.slug;
  else if (!slug || !verifyTrackingToken(slug, d, id, k)) {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const dealer = sess?.dealer || (await getDealer(slug));
  const order = dealer ? await getOrder(dealer.slug, d, id) : null;
  if (!dealer || !order) return NextResponse.json({ error: "Sipariş bulunamadı" }, { status: 404 });

  let pdf: Buffer;
  try {
    pdf = await generateOrderPdf(order, { name: dealer.name, phone: dealer.phone, website: dealer.website });
  } catch (e) {
    console.error("PDF üretilemedi:", e);
    return NextResponse.json({ error: "PDF üretilemedi" }, { status: 500 });
  }

  const rawName = `siparis_${order.orderId}_${order.customerName || ""}.pdf`;
  const ascii = asciiFileName(`siparis_${order.orderId}_${order.customerName || ""}`) || order.orderId;
  return new NextResponse(new Uint8Array(pdf), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${ascii}.pdf"; filename*=UTF-8''${encodeURIComponent(rawName)}`,
    },
  });
}
