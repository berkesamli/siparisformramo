import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { blobConfigured } from "@/lib/orders";
import {
  listCustomers,
  getCustomer,
  saveCustomer,
  deleteCustomer,
  sanitizeCustomer,
} from "@/lib/customers";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Müşteri defteri — yalnızca çalışanlar erişebilir.
export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  const customers = await listCustomers();
  return NextResponse.json({ customers, blob: blobConfigured() });
}

export async function POST(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  if (!blobConfigured()) {
    return NextResponse.json(
      { error: "Kalıcı depolama yapılandırılmadığı için müşteri kaydedilemiyor." },
      { status: 503 }
    );
  }

  const body = await req.json().catch(() => null);
  const c = sanitizeCustomer(body);
  if (!c.company && !c.firstName) {
    return NextResponse.json({ error: "Firma veya kişi adı gerekli." }, { status: 400 });
  }
  await saveCustomer(c);
  return NextResponse.json({ ok: true, customer: c });
}

export async function PUT(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  const body = await req.json().catch(() => null);
  const id = String(body?.id || "");
  if (!/^[A-Za-z0-9]{2,40}$/.test(id)) {
    return NextResponse.json({ error: "Geçersiz kayıt" }, { status: 400 });
  }
  const existing = await getCustomer(id);
  if (!existing) {
    return NextResponse.json({ error: "Müşteri bulunamadı" }, { status: 404 });
  }
  const c = sanitizeCustomer(body, existing);
  await saveCustomer(c);
  return NextResponse.json({ ok: true, customer: c });
}

export async function DELETE(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^[A-Za-z0-9]{2,40}$/.test(id)) {
    return NextResponse.json({ error: "Geçersiz kayıt" }, { status: 400 });
  }
  await deleteCustomer(id);
  return NextResponse.json({ ok: true });
}
