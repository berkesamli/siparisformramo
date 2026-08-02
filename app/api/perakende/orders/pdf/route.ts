import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getRetailOrder } from "@/lib/retail-orders";
import { generateRetailPdf } from "@/lib/retail-pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

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

  const pdf = await generateRetailPdf(order);
  const safeName = (order.customerName || order.orderId)
    .replace(/[^\p{L}\p{N} _-]/gu, "")
    .replace(/\s+/g, "_");
  return new NextResponse(new Uint8Array(pdf), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="olga_siparis_${safeName}.pdf"`,
    },
  });
}
