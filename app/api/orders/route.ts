import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import {
  sendOrderEmail,
  sendOrderWhatsApp,
  waLink,
  type OrderPayload,
  type OrderLine,
} from "@/lib/notify";
import {
  saveOrder,
  listOrders,
  lastNDateKeys,
  istanbulDateKey,
  sanitizeLines,
  computeTotals,
  nextOrderId,
  getDailyRates,
  saveDailyRates,
  type SavedOrder,
} from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";


// Sipariş listesi: ?range=today | ?range=week | ?date=YYYY-MM-DD
export async function GET(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const date = url.searchParams.get("date");
  const range = url.searchParams.get("range") || "today";

  let dateKeys: string[];
  if (date && /^\d{4}-\d{2}-\d{2}$/.test(date)) dateKeys = [date];
  else if (range === "week") dateKeys = lastNDateKeys(7);
  else dateKeys = lastNDateKeys(1);

  const orders = await listOrders(dateKeys);
  return NextResponse.json({ ok: true, orders });
}

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }

  const body = await req.json().catch(() => null);
  const lines = sanitizeLines(body?.lines);
  if (!lines.length) {
    return NextResponse.json({ ok: false, error: "Geçerli satır yok." }, { status: 400 });
  }

  const discountPct = Math.max(0, Number(body.discountPct) || 0);
  const vatApplied = !!body.vatApplied;
  const { gross, discount, vatAmount, net } = computeTotals(lines, discountPct, vatApplied);

  const now = new Date();
  const order: OrderPayload = {
    orderId: await nextOrderId(),
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
    dateStr: now.toLocaleString("tr-TR", { timeZone: "Europe/Istanbul" }),
  };

  // Kalıcı kayıt (sipariş takip ekranı için)
  const saved: SavedOrder = {
    orderId: order.orderId,
    dateKey: istanbulDateKey(now),
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
    status: "olusturuldu",
    employee: order.employee,
    customer: order.customer,
    // Müşteri defterinden seçildiyse cari takip için bağlanır
    customerId: String(body.customerId || "").slice(0, 40),
    payment: "bekliyor",
    paidAmount: 0,
    note: order.note,
    rate: order.rate,
    euroRate: order.euroRate,
    discountPct,
    vatApplied,
    lines,
    gross,
    discount,
    vatAmount,
    net,
    rows: Array.isArray(body.rows) ? body.rows.slice(0, 100) : undefined,
  };
  let stored = false;
  try {
    stored = await saveOrder(saved);
  } catch (err) {
    console.error("Sipariş Blob'a kaydedilemedi:", err);
  }

  // Günün kuru daha önce kaydedilmediyse bu siparişteki kuru günlük kur yap
  if (order.rate > 0) {
    try {
      const existing = await getDailyRates(saved.dateKey);
      if (!existing) {
        await saveDailyRates(saved.dateKey, {
          rate: order.rate,
          euroRate: order.euroRate,
          updatedAt: now.toISOString(),
          by: user.name,
        });
      }
    } catch (err) {
      console.error("Günlük kur kaydedilemedi:", err);
    }
  }

  let emailSent = false;
  let waSent = false;
  try {
    let pdf: Buffer | undefined;
    try {
      const { generateOrderPdf } = await import("@/lib/order-pdf");
      pdf = await generateOrderPdf(order);
    } catch (err) {
      console.error("PDF üretilemedi:", err);
    }
    emailSent = await sendOrderEmail(order, pdf);
  } catch (err) {
    console.error("E-posta gönderilemedi:", err);
  }
  try {
    waSent = await sendOrderWhatsApp(order);
  } catch (err) {
    console.error("WhatsApp gönderilemedi:", err);
  }

  // Müşteriye sipariş onay SMS'i — form işaretliyse ve müşteri defterden
  // seçilmişse. Telefonu yoksa veya SMS yapılandırılmadıysa sessizce atlanır;
  // SMS hatası sipariş kaydını asla geri döndürmez.
  let smsSent = false;
  let smsInfo = "";
  if (body.sendSms && saved.customerId) {
    try {
      const { smsConfigured, sendSms } = await import("@/lib/sms");
      if (smsConfigured()) {
        const { getCustomer } = await import("@/lib/customers");
        const c = await getCustomer(saved.customerId);
        if (c?.phone) {
          const { stripTurkish } = await import("@/lib/sms-format");
          const mesaj = stripTurkish(
            `Sayin musterimiz, ${order.orderId} numarali siparisiniz alinmistir. Tesekkur ederiz. Olga Cerceve`
          );
          const r = await sendSms([c.phone], mesaj);
          smsSent = r.ok;
          if (!r.ok) smsInfo = r.error || "SMS gönderilemedi.";

          const { saveSmsRecord, newSmsId, istanbulDateKey: gun } = await import(
            "@/lib/sms-log"
          );
          const t = new Date();
          await saveSmsRecord({
            id: newSmsId(t),
            dateKey: gun(t),
            createdAt: t.toISOString(),
            sender: user.name,
            message: mesaj,
            recipients: r.sent,
            segments: 1,
            credits: r.sent.length,
            ok: r.ok,
            jobId: r.jobId,
            error: r.error,
            iysfilter: "0",
          }).catch(() => {});
        } else {
          smsInfo = "Müşterinin kayıtlı telefonu yok.";
        }
      }
    } catch (err) {
      console.error("Sipariş SMS'i gönderilemedi:", err);
      smsInfo = "SMS gönderiminde hata oluştu.";
    }
  }

  return NextResponse.json({
    ok: true,
    orderId: order.orderId,
    dateKey: saved.dateKey,
    stored,
    emailSent,
    waSent,
    waLink: waSent ? undefined : waLink(order),
    smsSent,
    smsInfo,
    net,
  });
}
