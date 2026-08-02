import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { istanbulDateKey, lastNDateKeys, blobConfigured } from "@/lib/orders";
import {
  saveRetailOrder,
  getRetailOrder,
  listRetailOrders,
  nextRetailOrderId,
  type RetailItem,
  type SavedRetailOrder,
} from "@/lib/retail-orders";
import { sendRetailOrderEmail } from "@/lib/retail-notify";
import { generateRetailPdf } from "@/lib/retail-pdf";
import { RETAIL_STATUSES, type RetailStatus } from "@/data/perakende";

export const dynamic = "force-dynamic";

const r2 = (n: number) => Math.round((Number(n) || 0) * 100) / 100;
const s = (v: unknown, max = 200) => String(v ?? "").slice(0, max);

function sanitizeItems(raw: unknown): RetailItem[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((it: any): RetailItem => ({
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
      itemTotal: r2(it?.itemTotal),
    }))
    .filter((it) => it.artWidth > 0 && it.artHeight > 0);
}

export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const date = req.nextUrl.searchParams.get("date");
  const range = req.nextUrl.searchParams.get("range") || "today";
  let keys: string[];
  if (date && /^\d{4}-\d{2}-\d{2}$/.test(date)) keys = [date];
  else if (range === "week") keys = lastNDateKeys(7);
  else keys = [istanbulDateKey()];

  const orders = await listRetailOrders(keys);
  return NextResponse.json({ orders, blob: blobConfigured() });
}

export async function POST(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const body = await req.json().catch(() => null);
  if (!body) return NextResponse.json({ error: "Geçersiz istek" }, { status: 400 });

  const items = sanitizeItems(body.items);
  if (items.length === 0) {
    return NextResponse.json({ error: "Sipariş kalemi yok" }, { status: 400 });
  }
  const customerName = s(body.customerName, 120).trim();
  const customerPhone = s(body.customerPhone, 40).trim();
  if (!customerName || !customerPhone) {
    return NextResponse.json(
      { error: "Müşteri adı ve telefonu zorunludur" },
      { status: 400 }
    );
  }

  const gross = r2(items.reduce((sum, it) => sum + it.itemTotal, 0));
  const discount = Math.min(r2(body.discount), gross);
  const total = r2(gross - Math.max(0, discount));

  const now = new Date();
  const orderId = await nextRetailOrderId();
  const order: SavedRetailOrder = {
    orderId,
    dateKey: istanbulDateKey(now),
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
    status: "Beklemede",
    employee: user.name,
    customerName,
    customerPhone,
    customerEmail: s(body.customerEmail, 120).trim(),
    customerAddress: s(body.customerAddress, 240).trim(),
    usdRate: r2(body.usdRate),
    deliveryDate: s(body.deliveryDate, 20),
    notes: s(body.notes, 1000),
    items,
    gross,
    discount: Math.max(0, discount),
    total,
  };

  const saved = await saveRetailOrder(order);

  // Üretim PDF'i — e-posta ekinde gider, ayrıca listeden indirilebilir
  let pdf: Buffer | undefined;
  try {
    pdf = await generateRetailPdf(order);
  } catch {
    /* PDF hatası siparişi engellemesin */
  }

  let emailSent = false;
  try {
    emailSent = await sendRetailOrderEmail(order, pdf);
  } catch {
    /* e-posta hatası siparişi engellemesin */
  }

  return NextResponse.json({
    ok: true,
    orderId,
    dateKey: order.dateKey,
    saved,
    emailSent,
  });
}

export async function PATCH(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const d = req.nextUrl.searchParams.get("d") || "";
  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^PRK-[A-Z0-9-]+$/i.test(id)) {
    return NextResponse.json({ error: "Geçersiz sipariş" }, { status: 400 });
  }

  const body = await req.json().catch(() => null);
  const status = body?.status as RetailStatus;
  if (!RETAIL_STATUSES.includes(status)) {
    return NextResponse.json({ error: "Geçersiz durum" }, { status: 400 });
  }

  const order = await getRetailOrder(d, id);
  if (!order) return NextResponse.json({ error: "Sipariş bulunamadı" }, { status: 404 });

  order.status = status;
  order.updatedAt = new Date().toISOString();
  await saveRetailOrder(order);
  return NextResponse.json({ ok: true });
}
