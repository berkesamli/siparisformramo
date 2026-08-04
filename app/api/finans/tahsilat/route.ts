import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import {
  saveTahsilat,
  getTahsilat,
  deleteTahsilat,
  listTahsilatByMonths,
  listAllTahsilat,
  newTahsilatId,
  istanbulDateKey,
  type Tahsilat,
  type TahsilatYontem,
  type ParaBirimi,
} from "@/lib/tahsilat";
import { applyTahsilatDelta } from "@/lib/finans-ozet";
import { getOrder, saveOrder } from "@/lib/orders";
import { getRetailOrder, saveRetailOrder } from "@/lib/retail-orders";
import { kurus } from "@/lib/num";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const YONTEMLER: TahsilatYontem[] = [
  "nakit",
  "havale",
  "krediKarti",
  "cek",
  "senet",
  "diger",
];
const BIRIMLER: ParaBirimi[] = ["TL", "USD", "EUR"];

// Listeleme: filtresiz/ay bazlı liste finans yetkisi ister; belirli bir
// müşteri veya sipariş sorgusu tüm çalışanlara açıktır (cari sayfası staff'a
// açık olmaya devam ediyor).
export async function GET(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const musteri = url.searchParams.get("musteri");
  const siparis = url.searchParams.get("siparis");
  const ay = url.searchParams.get("ay");

  if (!musteri && !siparis && !isFinance(user.username)) {
    return NextResponse.json(
      { ok: false, error: "Bu listeye erişim yetkiniz yok." },
      { status: 403 }
    );
  }

  let records: Tahsilat[];
  if (ay && /^\d{4}-\d{2}$/.test(ay)) {
    records = await listTahsilatByMonths([ay]);
  } else {
    records = await listAllTahsilat();
  }
  if (musteri) records = records.filter((t) => t.customerId === musteri);
  if (siparis) records = records.filter((t) => t.orderId === siparis);
  return NextResponse.json({ ok: true, records });
}

/**
 * Bağlı siparişin paidAmount/payment alanını tahsilat toplamı kadar kaydırır.
 * delta pozitif = tahsilat eklendi, negatif = silindi.
 */
async function siparisTahsilatUygula(
  orderId: string,
  orderDateKey: string,
  delta: number
): Promise<void> {
  if (orderId.startsWith("PRK")) {
    const o = await getRetailOrder(orderDateKey, orderId);
    if (!o) return;
    const paid = Math.max(0, kurus((Number(o.paidAmount) || 0) + delta));
    o.paidAmount = Math.min(paid, o.total);
    o.payment =
      o.paidAmount <= 0 ? "bekliyor" : o.paidAmount >= o.total ? "odendi" : "kismi";
    o.updatedAt = new Date().toISOString();
    await saveRetailOrder(o);
  } else {
    const o = await getOrder(orderDateKey, orderId);
    if (!o) return;
    const paid = Math.max(0, kurus((Number(o.paidAmount) || 0) + delta));
    o.paidAmount = Math.min(paid, o.net);
    o.payment =
      o.paidAmount <= 0 ? "bekliyor" : o.paidAmount >= o.net ? "odendi" : "kismi";
    o.updatedAt = new Date().toISOString();
    await saveOrder(o);
  }
}

// Tahsilat girişi tüm çalışanlara açık — kasadaki personel kaydeder.
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as Record<string, unknown> | null;
  if (!body) {
    return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  }

  const amount = kurus(Number(body.amount) || 0);
  if (amount <= 0) {
    return NextResponse.json(
      { ok: false, error: "Tutar sıfırdan büyük olmalı." },
      { status: 400 }
    );
  }
  const method = String(body.method || "nakit") as TahsilatYontem;
  if (!YONTEMLER.includes(method)) {
    return NextResponse.json({ ok: false, error: "Geçersiz yöntem." }, { status: 400 });
  }
  const currency = String(body.currency || "TL") as ParaBirimi;
  if (!BIRIMLER.includes(currency)) {
    return NextResponse.json({ ok: false, error: "Geçersiz para birimi." }, { status: 400 });
  }
  const branch = body.branch === "istanbul" ? "istanbul" : "ankara";
  const customerName = String(body.customerName || "").trim().slice(0, 200);
  if (!customerName) {
    return NextResponse.json(
      { ok: false, error: "Müşteri adı gerekli." },
      { status: 400 }
    );
  }
  const dateKeyRaw = String(body.dateKey || "");
  const dateKey = /^\d{4}-\d{2}-\d{2}$/.test(dateKeyRaw)
    ? dateKeyRaw
    : istanbulDateKey();

  const now = new Date();
  const t: Tahsilat = {
    id: newTahsilatId(now),
    dateKey,
    createdAt: now.toISOString(),
    createdBy: user.name,
    branch,
    customerId: String(body.customerId || "").slice(0, 40) || undefined,
    customerName,
    orderId: String(body.orderId || "").slice(0, 40) || undefined,
    orderDateKey: String(body.orderDateKey || "").slice(0, 10) || undefined,
    amount,
    currency,
    method,
    tahsilEden: String(body.tahsilEden || "").trim().slice(0, 100) || undefined,
    note: String(body.note || "").trim().slice(0, 300) || undefined,
    kaynak: "panel",
  };

  const stored = await saveTahsilat(t);
  if (!stored) {
    return NextResponse.json(
      { ok: false, error: "Kalıcı depolama yapılandırılmadığı için kaydedilemedi." },
      { status: 503 }
    );
  }

  // Sipariş bağlantısı ve özet — hatalar tahsilat kaydını geri döndürmez;
  // özet zaten rebuild ile onarılabilir.
  if (t.orderId && t.orderDateKey && t.currency === "TL") {
    await siparisTahsilatUygula(t.orderId, t.orderDateKey, amount).catch(() => {});
  }
  await applyTahsilatDelta(t, 1).catch(() => {});

  return NextResponse.json({ ok: true, tahsilat: t });
}

// Silme: yalnızca finans yetkisi (yanlış girilen kaydın düzeltilmesi).
export async function DELETE(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isFinance(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const id = String(url.searchParams.get("id") || "");
  const ay = String(url.searchParams.get("ay") || "");
  if (!/^T-[\dA-Za-z-]+$/.test(id) || !/^\d{4}-\d{2}$/.test(ay)) {
    return NextResponse.json({ ok: false, error: "Geçersiz kayıt." }, { status: 400 });
  }
  const t = await getTahsilat(ay, id);
  if (!t) {
    return NextResponse.json({ ok: false, error: "Kayıt bulunamadı." }, { status: 404 });
  }
  await deleteTahsilat(ay, id);
  if (t.orderId && t.orderDateKey && t.currency === "TL") {
    await siparisTahsilatUygula(t.orderId, t.orderDateKey, -t.amount).catch(() => {});
  }
  await applyTahsilatDelta(t, -1).catch(() => {});
  return NextResponse.json({ ok: true });
}
