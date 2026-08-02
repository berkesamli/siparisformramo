import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { listAllOrders, orderBalance, type SavedOrder } from "@/lib/orders";
import { listAllRetailOrders, type SavedRetailOrder } from "@/lib/retail-orders";
import { getCustomer, customerTitle, normalizeCity } from "@/lib/customers";

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

  const [wholesale, retail] = await Promise.all([
    listAllOrders(),
    listAllRetailOrders(),
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

  const r2 = (n: number) => Math.round(n * 100) / 100;
  const summary = {
    orderCount: entries.length,
    totalAmount: r2(entries.reduce((s, e) => s + e.total, 0)),
    totalPaid: r2(entries.reduce((s, e) => s + e.paid, 0)),
    balance: r2(entries.reduce((s, e) => s + e.balance, 0)),
    lastOrderAt: entries[0]?.createdAt || null,
  };

  return NextResponse.json({ customer, entries, summary });
}
