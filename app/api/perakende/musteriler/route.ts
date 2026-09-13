// Perakende müşteri defteri API'si — etiket/toptan defterinden
// (/api/musteriler) ayrıdır. Perakende sihirbazındaki seçici ve
// /panel/perakende/musteriler yönetim sayfası bu ucu kullanır.
import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { blobConfigured } from "@/lib/orders";
import {
  listRetailCustomers,
  getRetailCustomer,
  saveRetailCustomer,
  deleteRetailCustomer,
  sanitizeRetailCustomer,
} from "@/lib/retail-customers";
import { eslesir } from "@/lib/search-norm";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  const q = (req.nextUrl.searchParams.get("q") || "").trim();
  let customers = await listRetailCustomers();
  if (q) {
    customers = customers.filter((c) =>
      eslesir(q, c.name, c.phone, c.email, c.address)
    );
  }
  return NextResponse.json({ ok: true, customers, blob: blobConfigured() });
}

export async function POST(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  const body = await req.json().catch(() => null);
  if (!body) return NextResponse.json({ error: "Geçersiz istek" }, { status: 400 });

  const existing = body.id ? await getRetailCustomer(String(body.id)) : null;
  const c = sanitizeRetailCustomer(body, existing || undefined);
  if (!c.name) {
    return NextResponse.json({ error: "Müşteri adı gerekli." }, { status: 400 });
  }
  const saved = await saveRetailCustomer(c);
  if (!saved) {
    return NextResponse.json(
      { error: "Kalıcı depo bağlı değil — müşteri kaydedilemedi." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, customer: c });
}

export async function DELETE(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  const id = req.nextUrl.searchParams.get("id") || "";
  if (!id) return NextResponse.json({ error: "Geçersiz id" }, { status: 400 });
  const okd = await deleteRetailCustomer(id);
  return NextResponse.json({ ok: okd });
}
