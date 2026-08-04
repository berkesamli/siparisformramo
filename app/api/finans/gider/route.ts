import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import {
  saveGider,
  getGider,
  deleteGider,
  listGiderByMonths,
  newGiderId,
  istanbulDateKey,
  type Gider,
  type GiderYontem,
} from "@/lib/gider";
import { applyGiderDelta } from "@/lib/finans-ozet";
import type { ParaBirimi } from "@/lib/tahsilat";
import { kurus } from "@/lib/num";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Giderler hassastır — tüm uçlar finans yetkisi ister.
const YONTEMLER: GiderYontem[] = ["nakit", "havale", "krediKarti", "cek", "diger"];
const BIRIMLER: ParaBirimi[] = ["TL", "USD", "EUR"];

async function yetkili() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isFinance(user.username)) return null;
  return user;
}

export async function GET(req: Request) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const ay = url.searchParams.get("ay");
  const sube = url.searchParams.get("sube");
  const month = ay && /^\d{4}-\d{2}$/.test(ay) ? ay : istanbulDateKey().slice(0, 7);
  let records = await listGiderByMonths([month]);
  if (sube === "ankara" || sube === "istanbul") {
    records = records.filter((g) => g.branch === sube);
  }
  return NextResponse.json({ ok: true, month, records });
}

export async function POST(req: Request) {
  const user = await yetkili();
  if (!user) {
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
  const method = String(body.method || "nakit") as GiderYontem;
  const currency = String(body.currency || "TL") as ParaBirimi;
  if (!YONTEMLER.includes(method) || !BIRIMLER.includes(currency)) {
    return NextResponse.json({ ok: false, error: "Geçersiz alan." }, { status: 400 });
  }
  const category = String(body.category || "").trim().slice(0, 60).toLocaleLowerCase("tr-TR");
  if (!category) {
    return NextResponse.json({ ok: false, error: "Kategori gerekli." }, { status: 400 });
  }
  const dateKeyRaw = String(body.dateKey || "");
  const now = new Date();
  const g: Gider = {
    id: newGiderId(now),
    dateKey: /^\d{4}-\d{2}-\d{2}$/.test(dateKeyRaw) ? dateKeyRaw : istanbulDateKey(),
    createdAt: now.toISOString(),
    createdBy: user.name,
    branch: body.branch === "istanbul" ? "istanbul" : "ankara",
    category,
    description: String(body.description || "").trim().slice(0, 300),
    amount,
    currency,
    method,
    supplier: String(body.supplier || "").trim().slice(0, 200) || undefined,
    personelId: String(body.personelId || "").slice(0, 40) || undefined,
    note: String(body.note || "").trim().slice(0, 300) || undefined,
    kaynak: "panel",
  };
  const stored = await saveGider(g);
  if (!stored) {
    return NextResponse.json(
      { ok: false, error: "Kalıcı depolama yapılandırılmadı." },
      { status: 503 }
    );
  }
  await applyGiderDelta(g, 1).catch(() => {});
  return NextResponse.json({ ok: true, gider: g });
}

export async function DELETE(req: Request) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const id = String(url.searchParams.get("id") || "");
  const ay = String(url.searchParams.get("ay") || "");
  if (!/^G-[\dA-Za-z-]+$/.test(id) || !/^\d{4}-\d{2}$/.test(ay)) {
    return NextResponse.json({ ok: false, error: "Geçersiz kayıt." }, { status: 400 });
  }
  const g = await getGider(ay, id);
  if (!g) {
    return NextResponse.json({ ok: false, error: "Kayıt bulunamadı." }, { status: 404 });
  }
  await deleteGider(ay, id);
  await applyGiderDelta(g, -1).catch(() => {});
  return NextResponse.json({ ok: true });
}
