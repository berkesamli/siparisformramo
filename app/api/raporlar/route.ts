import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { listAllOrders, orderBalance, blobConfigured } from "@/lib/orders";
import { listAllRetailOrders } from "@/lib/retail-orders";
import { findProfile } from "@/data/catalog";
import { isOwner } from "@/data/users";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const r2 = (n: number) => Math.round(n * 100) / 100;

// Satır adından profil kodunu çıkar: "GC065-1473 (Çerçeve Profili)" → GC065
function seriesOf(lineName: string): string | null {
  const code = String(lineName || "").trim().split(/[\s(]/)[0];
  if (!code) return null;
  const p = findProfile(code);
  return p ? p.series : null;
}

// Raporlar: /api/raporlar?ay=2026-08  (ay verilmezse tüm zamanlar)
export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  // Ciro/tahsilat verileri yalnızca firma sahiplerine döndürülür.
  if (!isOwner(user.username)) {
    return NextResponse.json({ error: "Bu rapora erişim yetkiniz yok." }, { status: 403 });
  }

  const ay = req.nextUrl.searchParams.get("ay") || "";
  const inRange = (dateKey: string) => (ay ? dateKey.startsWith(ay) : true);

  const [wholesale, retail] = await Promise.all([
    listAllOrders(),
    listAllRetailOrders(),
  ]);

  const w = wholesale.filter((o) => inRange(o.dateKey));
  const r = retail.filter((o) => inRange(o.dateKey));

  // ---- Ciro ----
  const toptanCiro = r2(w.reduce((s, o) => s + o.net, 0));
  const perakendeCiro = r2(r.reduce((s, o) => s + o.total, 0));

  // ---- Tahsilat / bakiye ----
  const paidOf = (net: number, payment?: string, paidAmount?: number) =>
    payment === "odendi" ? net : Number(paidAmount) || 0;
  const tahsilat = r2(
    w.reduce((s, o) => s + paidOf(o.net, o.payment, o.paidAmount), 0) +
      r.reduce((s, o) => s + paidOf(o.total, o.payment, o.paidAmount), 0)
  );
  const bakiye = r2(
    w.reduce((s, o) => s + orderBalance(o), 0) +
      r.reduce(
        (s, o) =>
          s + orderBalance({ net: o.total, payment: o.payment, paidAmount: o.paidAmount }),
        0
      )
  );

  // ---- Aylık dağılım (tüm zamanlar üzerinden) ----
  const byMonth = new Map<string, { toptan: number; perakende: number }>();
  wholesale.forEach((o) => {
    const k = o.dateKey.slice(0, 7);
    const cur = byMonth.get(k) || { toptan: 0, perakende: 0 };
    cur.toptan += o.net;
    byMonth.set(k, cur);
  });
  retail.forEach((o) => {
    const k = o.dateKey.slice(0, 7);
    const cur = byMonth.get(k) || { toptan: 0, perakende: 0 };
    cur.perakende += o.total;
    byMonth.set(k, cur);
  });
  const months = [...byMonth.entries()]
    .map(([month, v]) => ({
      month,
      toptan: r2(v.toptan),
      perakende: r2(v.perakende),
      toplam: r2(v.toptan + v.perakende),
    }))
    .sort((a, b) => b.month.localeCompare(a.month))
    .slice(0, 12);

  // ---- Müşteri bazlı satış ----
  const byCustomer = new Map<string, { total: number; count: number; balance: number }>();
  const addCustomer = (name: string, total: number, balance: number) => {
    const k = (name || "—").trim() || "—";
    const cur = byCustomer.get(k) || { total: 0, count: 0, balance: 0 };
    cur.total += total;
    cur.count += 1;
    cur.balance += balance;
    byCustomer.set(k, cur);
  };
  w.forEach((o) => addCustomer(o.customer, o.net, orderBalance(o)));
  r.forEach((o) =>
    addCustomer(
      o.customerName,
      o.total,
      orderBalance({ net: o.total, payment: o.payment, paidAmount: o.paidAmount })
    )
  );
  const customers = [...byCustomer.entries()]
    .map(([name, v]) => ({
      name,
      total: r2(v.total),
      count: v.count,
      balance: r2(v.balance),
    }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 15);

  // ---- En çok satan ürünler (toptan satırları) ----
  const byProduct = new Map<string, { total: number; count: number }>();
  const bySeries = new Map<string, number>();
  w.forEach((o) => {
    o.lines.forEach((l) => {
      const name = String(l.name || "").trim();
      if (!name) return;
      const cur = byProduct.get(name) || { total: 0, count: 0 };
      cur.total += l.lineTotal;
      cur.count += 1;
      byProduct.set(name, cur);

      const s = seriesOf(name);
      if (s) bySeries.set(s, (bySeries.get(s) || 0) + l.lineTotal);
    });
  });
  const products = [...byProduct.entries()]
    .map(([name, v]) => ({ name, total: r2(v.total), count: v.count }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 12);
  const series = [...bySeries.entries()]
    .map(([name, total]) => ({ name, total: r2(total) }))
    .sort((a, b) => b.total - a.total);

  // ---- Çalışan bazlı ----
  const byEmployee = new Map<string, { total: number; count: number }>();
  const addEmp = (name: string, total: number) => {
    const k = (name || "—").trim() || "—";
    const cur = byEmployee.get(k) || { total: 0, count: 0 };
    cur.total += total;
    cur.count += 1;
    byEmployee.set(k, cur);
  };
  w.forEach((o) => addEmp(o.employee, o.net));
  r.forEach((o) => addEmp(o.employee, o.total));
  const employees = [...byEmployee.entries()]
    .map(([name, v]) => ({ name, total: r2(v.total), count: v.count }))
    .sort((a, b) => b.total - a.total);

  return NextResponse.json({
    blob: blobConfigured(),
    range: ay || "tum",
    summary: {
      orderCount: w.length + r.length,
      toptanCount: w.length,
      perakendeCount: r.length,
      toptanCiro,
      perakendeCiro,
      toplamCiro: r2(toptanCiro + perakendeCiro),
      tahsilat,
      bakiye,
      ortalamaSepet:
        w.length + r.length > 0 ? r2((toptanCiro + perakendeCiro) / (w.length + r.length)) : 0,
    },
    months,
    customers,
    products,
    series,
    employees,
  });
}
