// İade — müşteriye para iadesi, ORİJİNAL SİPARİŞE BAĞLI NEGATİF tahsilat
// kaydı olarak işlenir: kasa hareketlerinde görünür, siparişin paidAmount'u
// düşer, cari bakiye kendiliğinden doğru kalır. Sipariş silinmez, durumu
// değişmez (gerekirse ayrıca "iptal" yapılır).
import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getOrder, saveOrder } from "@/lib/orders";
import { kurus } from "@/lib/num";
import {
  saveTahsilat,
  newTahsilatId,
  istanbulDateKey,
  type TahsilatYontem,
  TAHSILAT_YONTEM_LABELS,
} from "@/lib/tahsilat";
import { applyTahsilatDelta } from "@/lib/finans-ozet";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }

  const d = req.nextUrl.searchParams.get("d") || "";
  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !id) {
    return NextResponse.json({ ok: false, error: "Geçersiz sipariş." }, { status: 400 });
  }

  const body = await req.json().catch(() => null);
  const tutar = kurus(Number(body?.amount) || 0);
  if (!(tutar > 0)) {
    return NextResponse.json({ ok: false, error: "Geçerli bir iade tutarı girin." }, { status: 400 });
  }
  const yontem: TahsilatYontem =
    body?.method && body.method in TAHSILAT_YONTEM_LABELS ? body.method : "nakit";

  const order = await getOrder(d, id);
  if (!order) {
    return NextResponse.json({ ok: false, error: "Sipariş bulunamadı." }, { status: 404 });
  }

  const odenen =
    order.payment === "odendi" ? order.net : Number(order.paidAmount) || 0;
  if (tutar > odenen + 0.01) {
    return NextResponse.json(
      { ok: false, error: `İade, ödenenden fazla olamaz (ödenen ₺${odenen}).` },
      { status: 400 }
    );
  }

  // 1) Negatif tahsilat — kasadan çıkışın kalıcı kaydı
  const t = {
    id: newTahsilatId(),
    dateKey: istanbulDateKey(),
    createdAt: new Date().toISOString(),
    createdBy: user.name,
    branch: (order.branch || "ankara") as "ankara" | "istanbul",
    customerId: order.customerId || undefined,
    customerName: order.customer,
    orderId: order.orderId,
    orderDateKey: order.dateKey,
    amount: -tutar,
    currency: "TL" as const,
    method: yontem,
    note: `İade${body?.note ? ` — ${String(body.note).slice(0, 200)}` : ""}`,
    kaynak: "panel" as const,
  };
  const kayitOldu = await saveTahsilat(t);
  if (!kayitOldu) {
    return NextResponse.json(
      { ok: false, error: "İade kaydı yazılamadı (depo bağlı değil)." },
      { status: 503 }
    );
  }
  try {
    await applyTahsilatDelta(t, 1);
  } catch (err) {
    console.error("İade kasa özetine işlenemedi:", err);
  }

  // 2) Siparişin ödenen tutarı düşer, ödeme durumu yeniden hesaplanır
  const yeniOdenen = kurus(Math.max(0, odenen - tutar));
  order.paidAmount = yeniOdenen;
  order.payment =
    yeniOdenen <= 0 ? "bekliyor" : yeniOdenen + 0.01 >= order.net ? "odendi" : "kismi";
  order.updatedAt = new Date().toISOString();
  try {
    await saveOrder(order);
  } catch (err) {
    console.error("İade sonrası sipariş güncellenemedi:", err);
  }

  return NextResponse.json({
    ok: true,
    tahsilatId: t.id,
    paidAmount: order.paidAmount,
    payment: order.payment,
  });
}
