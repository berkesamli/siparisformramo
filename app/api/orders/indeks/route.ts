// Aylık sipariş indeksini baştan kurar (yalnız firma sahipleri).
// Normalde gerekmez — indeks her kayıtta kendiliğinden güncellenir ve boşken
// ilk aramada kendini onarır. Bu uç, indeksin bozulduğundan şüphelenildiğinde
// elle tazelemek içindir: POST /api/orders/indeks
import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import { listAllOrders, rebuildOrderIndexFrom, blobConfigured } from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 120;

export async function POST() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  if (!blobConfigured()) {
    return NextResponse.json(
      { ok: false, error: "Blob deposu yapılandırılmamış." },
      { status: 503 }
    );
  }
  const orders = await listAllOrders();
  const ay = await rebuildOrderIndexFrom(orders);
  return NextResponse.json({ ok: true, siparis: orders.length, ay });
}
