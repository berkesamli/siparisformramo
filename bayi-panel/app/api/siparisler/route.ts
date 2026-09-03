import { NextRequest, NextResponse } from "next/server";
import { getDealerSession } from "@/lib/auth";
import { dealerCanOrder } from "@/lib/dealers";
import { blobConfigured, istanbulDateKey, lastNDateKeys } from "@/lib/store";
import {
  saveOrder,
  getOrder,
  listOrders,
  listAllOrders,
  nextOrderId,
  trackingPath,
  type OrderItem,
  type SavedOrder,
} from "@/lib/orders";
import { generateOrderPdf } from "@/lib/pdf";
import { sendOrderEmail } from "@/lib/notify";
import { ORDER_STATUSES, PAYMENT_LABELS, type OrderStatus, type PaymentStatus } from "@/data/pricing";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const r2 = (n: number) => Math.round((Number(n) || 0) * 100) / 100;
const s = (v: unknown, max = 200) => String(v ?? "").slice(0, max);

function sanitizeItems(raw: unknown): OrderItem[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((it: any): OrderItem => ({
      artWidth: Number(it?.artWidth) || 0,
      artWidthUnit: it?.artWidthUnit === "mm" ? "mm" : "cm",
      artHeight: Number(it?.artHeight) || 0,
      artHeightUnit: it?.artHeightUnit === "mm" ? "mm" : "cm",
      frameCode: s(it?.frameCode, 60),
      framePriceTL: r2(it?.framePriceTL),
      manualPrice: Boolean(it?.manualPrice),
      matType: s(it?.matType, 60),
      matCode: s(it?.matCode, 20),
      matColor: s(it?.matColor, 20),
      matColorHex: s(it?.matColorHex, 20),
      doubleMat: Boolean(it?.doubleMat),
      innerMatType: s(it?.innerMatType, 60),
      innerMatColor: s(it?.innerMatColor, 20),
      innerMatColorHex: s(it?.innerMatColorHex, 20),
      altMontaj: s(it?.altMontaj, 10),
      zeminEnabled: Boolean(it?.zeminEnabled),
      zeminType: s(it?.zeminType, 60),
      zeminColor: s(it?.zeminColor, 20),
      zeminColorHex: s(it?.zeminColorHex, 20),
      matTop: Number(it?.matTop) || 0,
      matRight: Number(it?.matRight) || 0,
      matBottom: Number(it?.matBottom) || 0,
      matLeft: Number(it?.matLeft) || 0,
      glassType: s(it?.glassType, 60),
      printType: s(it?.printType, 60),
      frameCost: r2(it?.frameCost),
      matCost: r2(it?.matCost),
      glassCost: r2(it?.glassCost),
      printCost: r2(it?.printCost),
      laborCost: r2(it?.laborCost),
      itemTotal: r2(it?.itemTotal),
    }))
    .filter((it) => it.artWidth > 0 && it.artHeight > 0);
}

function withTrack(o: SavedOrder, origin: string) {
  return { ...o, trackUrl: origin + trackingPath(o) };
}

export async function GET(req: NextRequest) {
  const sess = await getDealerSession();
  if (!sess) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });

  const slug = sess.dealer.slug;
  const date = req.nextUrl.searchParams.get("date");
  const range = req.nextUrl.searchParams.get("range") || "today";
  let orders: SavedOrder[];
  if (date && /^\d{4}-\d{2}-\d{2}$/.test(date)) orders = await listOrders(slug, [date]);
  else if (range === "week") orders = await listOrders(slug, lastNDateKeys(7));
  else if (range === "month") orders = await listOrders(slug, lastNDateKeys(30));
  else if (range === "all") orders = await listAllOrders(slug);
  else orders = await listOrders(slug, [istanbulDateKey()]);

  const origin = req.nextUrl.origin;
  return NextResponse.json({ orders: orders.map((o) => withTrack(o, origin)), blob: blobConfigured() });
}

export async function POST(req: NextRequest) {
  const sess = await getDealerSession();
  if (!sess) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  const can = dealerCanOrder(sess.dealer);
  if (!can.ok) return NextResponse.json({ error: can.reason }, { status: 403 });
  if (!blobConfigured()) {
    return NextResponse.json({ error: "Kalıcı depolama (Blob) yapılandırılmamış; sipariş kaydedilemez." }, { status: 503 });
  }

  const body = await req.json().catch(() => null);
  if (!body) return NextResponse.json({ error: "Geçersiz istek" }, { status: 400 });

  const items = sanitizeItems(body.items);
  if (items.length === 0) return NextResponse.json({ error: "Sipariş kalemi yok" }, { status: 400 });
  const customerName = s(body.customerName, 120).trim();
  const customerPhone = s(body.customerPhone, 40).trim();
  if (!customerName || !customerPhone) {
    return NextResponse.json({ error: "Müşteri adı ve telefonu zorunludur" }, { status: 400 });
  }

  const gross = r2(items.reduce((sum, it) => sum + it.itemTotal, 0));
  const discount = Math.max(0, Math.min(r2(body.discount), gross));
  const total = r2(gross - discount);

  const now = new Date();
  const slug = sess.dealer.slug;
  const orderId = await nextOrderId(slug);
  const order: SavedOrder = {
    orderId,
    dealerSlug: slug,
    dateKey: istanbulDateKey(now),
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
    status: "Beklemede",
    createdBy: sess.user.name,
    customerName,
    customerPhone,
    customerEmail: s(body.customerEmail, 120).trim(),
    customerAddress: s(body.customerAddress, 240).trim(),
    payment: "bekliyor",
    paidAmount: 0,
    usdRate: r2(body.usdRate),
    deliveryDate: s(body.deliveryDate, 20),
    notes: s(body.notes, 1000),
    items,
    gross,
    discount,
    total,
  };

  const saved = await saveOrder(order);
  const trackUrl = req.nextUrl.origin + trackingPath(order);
  const brand = { name: sess.dealer.name, phone: sess.dealer.phone, website: sess.dealer.website };

  let emailSent = false;
  if (order.customerEmail) {
    try {
      let pdf: Buffer | undefined;
      try {
        pdf = await generateOrderPdf(order, brand);
      } catch {
        /* PDF hatası e-postayı engellemesin */
      }
      emailSent = await sendOrderEmail(order, brand, trackUrl, pdf);
    } catch {
      /* e-posta hatası siparişi engellemesin */
    }
  }

  return NextResponse.json({ ok: true, orderId, dateKey: order.dateKey, saved, emailSent, trackUrl });
}

export async function PATCH(req: NextRequest) {
  const sess = await getDealerSession();
  if (!sess) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });

  const d = req.nextUrl.searchParams.get("d") || "";
  const id = req.nextUrl.searchParams.get("id") || "";
  const body = await req.json().catch(() => null);
  const order = await getOrder(sess.dealer.slug, d, id);
  if (!order) return NextResponse.json({ error: "Sipariş bulunamadı" }, { status: 404 });

  if (body?.status !== undefined) {
    const status = body.status as OrderStatus;
    if (!ORDER_STATUSES.includes(status)) {
      return NextResponse.json({ error: "Geçersiz durum" }, { status: 400 });
    }
    order.status = status;
  }
  if (body?.payment !== undefined) {
    const payment = String(body.payment) as PaymentStatus;
    if (!(payment in PAYMENT_LABELS)) {
      return NextResponse.json({ error: "Geçersiz ödeme durumu" }, { status: 400 });
    }
    order.payment = payment;
    if (payment === "odendi") order.paidAmount = order.total;
    else if (payment === "bekliyor") order.paidAmount = 0;
  }
  if (body?.paidAmount !== undefined) {
    const paid = Math.max(0, r2(body.paidAmount));
    order.paidAmount = paid;
    if (paid <= 0) order.payment = "bekliyor";
    else if (paid + 0.01 >= order.total) {
      order.payment = "odendi";
      order.paidAmount = order.total;
    } else order.payment = "kismi";
  }
  if (body?.deliveryDate !== undefined) order.deliveryDate = s(body.deliveryDate, 20);
  if (body?.notes !== undefined) order.notes = s(body.notes, 1000);

  order.updatedAt = new Date().toISOString();
  await saveOrder(order);
  return NextResponse.json({ ok: true, status: order.status, payment: order.payment, paidAmount: order.paidAmount });
}
