import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import {
  sendOrderEmail,
  sendOrderWhatsApp,
  waLink,
  type OrderPayload,
  type OrderLine,
} from "@/lib/notify";

export const runtime = "nodejs";

function makeOrderId(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  const stamp = `${String(d.getFullYear()).slice(2)}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
  const rand = Math.random().toString(36).slice(2, 5).toUpperCase();
  return `OLG-${stamp}-${rand}`;
}

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }

  const body = await req.json().catch(() => null);
  if (!body || !Array.isArray(body.lines) || body.lines.length === 0) {
    return NextResponse.json({ ok: false, error: "Satır verisi boş." }, { status: 400 });
  }

  const lines: OrderLine[] = body.lines
    .map((l: any) => ({
      name: String(l.name || "").slice(0, 200),
      unitText: String(l.unitText || "").slice(0, 200),
      unitPriceTL: Number(l.unitPriceTL) || 0,
      lineTotal: Number(l.lineTotal) || 0,
    }))
    .filter((l: OrderLine) => l.name);

  if (!lines.length) {
    return NextResponse.json({ ok: false, error: "Geçerli satır yok." }, { status: 400 });
  }

  const r2 = (n: number) => Math.round(n * 100) / 100;
  const gross = r2(lines.reduce((s, l) => s + l.lineTotal, 0));
  const discountPct = Math.max(0, Number(body.discountPct) || 0);
  const discount = r2(gross * (discountPct / 100));
  const afterDiscount = r2(Math.max(0, gross - discount));
  const vatApplied = !!body.vatApplied;
  const vatAmount = vatApplied ? r2(afterDiscount * 0.2) : 0;
  const net = r2(afterDiscount + vatAmount);

  const order: OrderPayload = {
    orderId: makeOrderId(),
    employee: user.name,
    customer: String(body.customer || "").slice(0, 200),
    note: String(body.note || "").slice(0, 500),
    rate: Number(body.rate) || 0,
    euroRate: Number(body.euroRate) || 0,
    discountPct,
    vatApplied,
    lines,
    gross,
    discount,
    vatAmount,
    net,
    dateStr: new Date().toLocaleString("tr-TR", { timeZone: "Europe/Istanbul" }),
  };

  let emailSent = false;
  let waSent = false;
  try {
    emailSent = await sendOrderEmail(order);
  } catch (err) {
    console.error("E-posta gönderilemedi:", err);
  }
  try {
    waSent = await sendOrderWhatsApp(order);
  } catch (err) {
    console.error("WhatsApp gönderilemedi:", err);
  }

  return NextResponse.json({
    ok: true,
    orderId: order.orderId,
    emailSent,
    waSent,
    waLink: waSent ? undefined : waLink(order),
    net,
  });
}
