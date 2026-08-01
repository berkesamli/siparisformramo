import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getOrder, STATUS_LABELS } from "@/lib/orders";
import { generateOrderPdf } from "@/lib/order-pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Sipariş fişini PDF olarak indirir: /api/orders/pdf?d=YYYY-MM-DD&id=OLG-...
export async function GET(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const dateKey = url.searchParams.get("d") || "";
  const orderId = url.searchParams.get("id") || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey) || !/^OLG-[A-Z0-9-]+$/i.test(orderId)) {
    return NextResponse.json({ ok: false, error: "Geçersiz parametre." }, { status: 400 });
  }
  const order = await getOrder(dateKey, orderId);
  if (!order) {
    return NextResponse.json({ ok: false, error: "Sipariş bulunamadı." }, { status: 404 });
  }

  const pdf = await generateOrderPdf({
    orderId: order.orderId,
    dateStr: new Date(order.createdAt).toLocaleString("tr-TR", {
      timeZone: "Europe/Istanbul",
    }),
    status: STATUS_LABELS[order.status],
    employee: order.employee,
    customer: order.customer,
    note: order.note,
    discountPct: order.discountPct,
    vatApplied: order.vatApplied,
    lines: order.lines,
    gross: order.gross,
    discount: order.discount,
    vatAmount: order.vatAmount,
    net: order.net,
  });

  return new NextResponse(new Uint8Array(pdf), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="Siparis_${order.orderId}.pdf"`,
    },
  });
}
