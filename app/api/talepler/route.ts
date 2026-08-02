import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { blobConfigured } from "@/lib/orders";
import {
  listRequests,
  getRequest,
  saveRequest,
  newRequest,
  sanitizeRequestLines,
  REQUEST_LABELS,
  type RequestStatus,
} from "@/lib/requests";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// GET: çalışan tüm talepleri, bayi yalnızca kendi taleplerini görür
export async function GET() {
  const user = await getSessionUser();
  if (!user) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });

  const all = await listRequests();
  const requests =
    user.role === "staff" ? all : all.filter((r) => r.username === user.username);
  return NextResponse.json({ requests, blob: blobConfigured() });
}

// POST: bayi yeni talep gönderir
export async function POST(req: NextRequest) {
  const user = await getSessionUser();
  if (!user) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  if (!blobConfigured()) {
    return NextResponse.json(
      { error: "Kalıcı depolama yapılandırılmadığı için talep kaydedilemiyor." },
      { status: 503 }
    );
  }

  const body = await req.json().catch(() => null);
  const lines = sanitizeRequestLines(body?.lines);
  if (lines.length === 0) {
    return NextResponse.json({ error: "En az bir ürün satırı gerekli." }, { status: 400 });
  }

  const request = newRequest({
    username: user.username,
    customer: String(body?.customer || user.name).slice(0, 160),
    phone: String(body?.phone || "").slice(0, 40),
    note: String(body?.note || "").slice(0, 500),
    lines,
  });
  await saveRequest(request);
  return NextResponse.json({ ok: true, request });
}

// PATCH: çalışan talebi onaylar/reddeder — ?d=&id=
export async function PATCH(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }
  const d = req.nextUrl.searchParams.get("d") || "";
  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^T[A-Z0-9]+$/i.test(id)) {
    return NextResponse.json({ error: "Geçersiz talep" }, { status: 400 });
  }

  const body = await req.json().catch(() => null);
  const status = String(body?.status || "") as RequestStatus;
  if (!(status in REQUEST_LABELS)) {
    return NextResponse.json({ error: "Geçersiz durum" }, { status: 400 });
  }

  const request = await getRequest(d, id);
  if (!request) return NextResponse.json({ error: "Talep bulunamadı" }, { status: 404 });

  request.status = status;
  request.handledBy = user.name;
  if (body?.orderId) request.orderId = String(body.orderId).slice(0, 40);
  request.updatedAt = new Date().toISOString();
  await saveRequest(request);
  return NextResponse.json({ ok: true, request });
}
