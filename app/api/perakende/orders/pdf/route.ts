import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getRetailOrder } from "@/lib/retail-orders";
import { generateRetailPdf } from "@/lib/retail-pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// HTTP başlıkları yalnızca ASCII taşıyabilir; Türkçe karakterli müşteri adları
// dosya adında başlığı geçersiz kılıyordu. ASCII karşılığı filename'e, gerçek
// ad ise RFC 5987 filename* alanına yazılır.
const TR_ASCII: Record<string, string> = {
  ç: "c", Ç: "C", ğ: "g", Ğ: "G", ı: "i", İ: "I",
  ö: "o", Ö: "O", ş: "s", Ş: "S", ü: "u", Ü: "U",
};

function asciiFileName(s: string): string {
  return s
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_ASCII[c] || c)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Za-z0-9 _-]/g, "")
    .trim()
    .replace(/\s+/g, "_");
}

// Perakende sipariş üretim PDF'i indirme (çalışanlara özel)
export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const d = req.nextUrl.searchParams.get("d") || "";
  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^PRK-[A-Z0-9-]+$/i.test(id)) {
    return NextResponse.json({ error: "Geçersiz sipariş" }, { status: 400 });
  }

  const order = await getRetailOrder(d, id);
  if (!order) {
    return NextResponse.json({ error: "Sipariş bulunamadı" }, { status: 404 });
  }

  let pdf: Buffer;
  try {
    pdf = await generateRetailPdf(order);
  } catch (e) {
    console.error("Perakende PDF üretilemedi:", e);
    return NextResponse.json(
      { error: "PDF üretilemedi: " + (e instanceof Error ? e.message : "bilinmeyen hata") },
      { status: 500 }
    );
  }

  const rawName = `olga_siparis_${order.customerName || order.orderId}.pdf`;
  const ascii =
    asciiFileName(order.customerName || "") || order.orderId.replace(/[^A-Za-z0-9-]/g, "");

  return new NextResponse(new Uint8Array(pdf), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition":
        `attachment; filename="olga_siparis_${ascii}.pdf"; ` +
        `filename*=UTF-8''${encodeURIComponent(rawName)}`,
    },
  });
}
