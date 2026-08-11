import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { listAllOrders, orderBalance, type SavedOrder } from "@/lib/orders";
import { listAllRetailOrders, type SavedRetailOrder } from "@/lib/retail-orders";
import { getCustomer, customerTitle, normalizeCity } from "@/lib/customers";
import { listAllTahsilat, type Tahsilat } from "@/lib/tahsilat";
import { getAcilisBakiye } from "@/lib/acilis-bakiye";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Müşteri kartındaki isim/firma ile sipariş üzerindeki müşteri adını
// eşleştirmek için sadeleştirme (eski siparişlerde customerId yok).
function key(s: string): string {
  return normalizeCity(s).replace(/[^a-z0-9]/g, "");
}

export interface CariEntry {
  kind: "toptan" | "perakende";
  orderId: string;
  dateKey: string;
  createdAt: string;
  status: string;
  total: number;
  paid: number;
  balance: number;
}

// Bir müşterinin sipariş geçmişi + cari özeti: /api/musteriler/cari?id=C123
export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^[A-Za-z0-9]{2,40}$/.test(id)) {
    return NextResponse.json({ error: "Geçersiz müşteri" }, { status: 400 });
  }
  const customer = await getCustomer(id);
  if (!customer) {
    return NextResponse.json({ error: "Müşteri bulunamadı" }, { status: 404 });
  }

  // Tek istekte üç tarama — başka hiçbir uçta tam tarama yapılmaz.
  const [wholesale, retail, tahsilatlar, acilis] = await Promise.all([
    listAllOrders(),
    listAllRetailOrders(),
    listAllTahsilat(),
    getAcilisBakiye(id),
  ]);

  // Eşleşme: önce customerId, yoksa isim benzerliği
  const names = new Set(
    [customerTitle(customer), customer.company, `${customer.firstName} ${customer.lastName}`]
      .map(key)
      .filter((k) => k.length > 2)
  );
  const matchesName = (n: string) => {
    const k = key(n);
    if (!k) return false;
    for (const want of names) {
      if (k === want || k.includes(want) || want.includes(k)) return true;
    }
    return false;
  };

  const entries: CariEntry[] = [];

  wholesale.forEach((o: SavedOrder) => {
    const mine = o.customerId ? o.customerId === id : matchesName(o.customer);
    if (!mine) return;
    // İptal edilen sipariş cari bakiyeye ve hareketlere girmez
    if (o.status === "iptal") return;
    const paid = o.payment === "odendi" ? o.net : Number(o.paidAmount) || 0;
    entries.push({
      kind: "toptan",
      orderId: o.orderId,
      dateKey: o.dateKey,
      createdAt: o.createdAt,
      status: o.status,
      total: o.net,
      paid,
      balance: orderBalance(o),
    });
  });

  retail.forEach((o: SavedRetailOrder) => {
    const mine = o.customerId ? o.customerId === id : matchesName(o.customerName);
    if (!mine) return;
    if (o.status === "İptal") return;
    const paid = o.payment === "odendi" ? o.total : Number(o.paidAmount) || 0;
    entries.push({
      kind: "perakende",
      orderId: o.orderId,
      dateKey: o.dateKey,
      createdAt: o.createdAt,
      status: o.status,
      total: o.total,
      paid,
      balance: orderBalance({
        net: o.total,
        payment: o.payment,
        paidAmount: o.paidAmount,
      }),
    });
  });

  entries.sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  // Tahsilat hareketleri: customerId eşleşmesi veya isim benzerliği.
  // Açılış devri varsa devir tarihinden önceki hareketler Excel'de kapanmıştır,
  // bakiyeye tekrar katılmaz.
  const asOf = acilis?.asOf || "";
  const movements = tahsilatlar.filter((t: Tahsilat) => {
    const mine = t.customerId ? t.customerId === id : matchesName(t.customerName);
    if (!mine) return false;
    return !asOf || t.dateKey >= asOf;
  });

  const r2 = (n: number) => Math.round(n * 100) / 100;
  const entriesInScope = asOf
    ? entries.filter((e) => e.dateKey >= asOf)
    : entries;
  const totalAmount = r2(entriesInScope.reduce((s, e) => s + e.total, 0));
  // Tahsil edilen: gerçek tahsilat kayıtları (TL) esas alınır; hiç kayıt yoksa
  // eski paidAmount toplamına düşülür (geçiş dönemi uyumu).
  const tahsilatToplam = r2(
    movements.filter((t) => t.currency === "TL").reduce((s, t) => s + t.amount, 0)
  );
  const eskiPaid = r2(entriesInScope.reduce((s, e) => s + e.paid, 0));
  const totalPaid = movements.length ? tahsilatToplam : eskiPaid;
  const opening = acilis ? r2(acilis.amount) : 0;
  const balance = r2(opening + totalAmount - totalPaid);

  const summary = {
    orderCount: entries.length,
    totalAmount,
    totalPaid,
    openingBalance: opening,
    openingAsOf: acilis?.asOf || null,
    balance,
    lastOrderAt: entries[0]?.createdAt || null,
  };

  return NextResponse.json({ customer, entries, movements, summary });
}
