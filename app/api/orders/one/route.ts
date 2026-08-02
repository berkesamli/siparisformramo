import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import {
  getOrder,
  saveOrder,
  sanitizeLines,
  computeTotals,
  STATUS_LABELS,
  PAYMENT_LABELS,
  type OrderStatus,
  type PaymentStatus,
} from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function params(req: Request) {
  const url = new URL(req.url);
  return {
    dateKey: url.searchParams.get("d") || "",
    orderId: url.searchParams.get("id") || "",
  };
}

const valid = (d: string, id: string) =>
  /^\d{4}-\d{2}-\d{2}$/.test(d) && /^OLG-[A-Z0-9-]+$/i.test(id);

// Tek sipariş getir
export async function GET(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const { dateKey, orderId } = params(req);
  if (!valid(dateKey, orderId)) {
    return NextResponse.json({ ok: false, error: "Geçersiz parametre." }, { status: 400 });
  }
  const order = await getOrder(dateKey, orderId);
  if (!order) {
    return NextResponse.json({ ok: false, error: "Sipariş bulunamadı." }, { status: 404 });
  }
  return NextResponse.json({ ok: true, order });
}

// Durum güncelle: { status }
export async function PATCH(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const { dateKey, orderId } = params(req);
  if (!valid(dateKey, orderId)) {
    return NextResponse.json({ ok: false, error: "Geçersiz parametre." }, { status: 400 });
  }
  const body = await req.json().catch(() => null);
  const order = await getOrder(dateKey, orderId);
  if (!order) {
    return NextResponse.json({ ok: false, error: "Sipariş bulunamadı." }, { status: 404 });
  }

  // Durum güncelleme
  if (body?.status !== undefined) {
    const status = String(body.status) as OrderStatus;
    if (!(status in STATUS_LABELS)) {
      return NextResponse.json({ ok: false, error: "Geçersiz durum." }, { status: 400 });
    }
    order.status = status;
  }

  // Ödeme (cari) güncelleme
  if (body?.payment !== undefined) {
    const payment = String(body.payment) as PaymentStatus;
    if (!(payment in PAYMENT_LABELS)) {
      return NextResponse.json({ ok: false, error: "Geçersiz ödeme durumu." }, { status: 400 });
    }
    order.payment = payment;
    if (payment === "odendi") order.paidAmount = order.net;
    else if (payment === "bekliyor") order.paidAmount = 0;
  }
  if (body?.paidAmount !== undefined) {
    const paid = Math.max(0, Math.round((Number(body.paidAmount) || 0) * 100) / 100);
    order.paidAmount = paid;
    // Tutara göre durumu kendiliğinden ayarla
    if (paid <= 0) order.payment = "bekliyor";
    else if (paid + 0.01 >= order.net) {
      order.payment = "odendi";
      order.paidAmount = order.net;
    } else order.payment = "kismi";
  }

  order.updatedAt = new Date().toISOString();
  await saveOrder(order);
  return NextResponse.json({
    ok: true,
    status: order.status,
    payment: order.payment,
    paidAmount: order.paidAmount,
  });
}

// Siparişi düzenle: satırlar + üst bilgiler yeniden kaydedilir
export async function PUT(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const { dateKey, orderId } = params(req);
  if (!valid(dateKey, orderId)) {
    return NextResponse.json({ ok: false, error: "Geçersiz parametre." }, { status: 400 });
  }
  const existing = await getOrder(dateKey, orderId);
  if (!existing) {
    return NextResponse.json({ ok: false, error: "Sipariş bulunamadı." }, { status: 404 });
  }

  const body = await req.json().catch(() => null);
  const lines = sanitizeLines(body?.lines);
  if (!lines.length) {
    return NextResponse.json({ ok: false, error: "Geçerli satır yok." }, { status: 400 });
  }
  const discountPct = Math.max(0, Number(body.discountPct) || 0);
  const vatApplied = !!body.vatApplied;
  const { gross, discount, vatAmount, net } = computeTotals(lines, discountPct, vatApplied);

  const updated = {
    ...existing,
    customer: String(body.customer || existing.customer).slice(0, 200),
    note: String(body.note ?? existing.note).slice(0, 500),
    rate: Number(body.rate) || 0,
    euroRate: Number(body.euroRate) || 0,
    discountPct,
    vatApplied,
    lines,
    gross,
    discount,
    vatAmount,
    net,
    rows: Array.isArray(body.rows) ? body.rows.slice(0, 100) : existing.rows,
    updatedAt: new Date().toISOString(),
  };
  await saveOrder(updated);
  return NextResponse.json({ ok: true, orderId, net });
}
