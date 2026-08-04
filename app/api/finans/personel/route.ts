import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import {
  savePersonel,
  getPersonel,
  listPersonel,
  newPersonelId,
  type Personel,
} from "@/lib/personel";
import { listGiderByMonths, istanbulDateKey } from "@/lib/gider";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

async function yetkili() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isFinance(user.username)) return null;
  return user;
}

// Personel listesi + seçilen ayın maaş/avans/prim ödemeleri (tek istekte).
export async function GET(req: Request) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const ayParam = url.searchParams.get("ay");
  const ay =
    ayParam && /^\d{4}-\d{2}$/.test(ayParam)
      ? ayParam
      : istanbulDateKey().slice(0, 7);
  const [personel, giderler] = await Promise.all([
    listPersonel(),
    listGiderByMonths([ay]),
  ]);
  const odemeler = giderler.filter(
    (g) => g.personelId || ["maaş", "avans", "prim"].includes(g.category)
  );
  return NextResponse.json({ ok: true, ay, personel, odemeler });
}

export async function POST(req: Request) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as Record<string, unknown> | null;
  const name = String(body?.name || "").trim().slice(0, 120);
  if (!name) {
    return NextResponse.json({ ok: false, error: "İsim gerekli." }, { status: 400 });
  }
  const id = String(body?.id || "").slice(0, 40);
  const now = new Date().toISOString();
  const mevcut = id ? await getPersonel(id) : null;
  const p: Personel = {
    id: mevcut?.id || newPersonelId(),
    name,
    branch: body?.branch === "istanbul" ? "istanbul" : "ankara",
    startDate: String(body?.startDate || "").slice(0, 10) || undefined,
    endDate: String(body?.endDate || "").slice(0, 10) || undefined,
    salary: Number(body?.salary) || undefined,
    note: String(body?.note || "").trim().slice(0, 300) || undefined,
    createdAt: mevcut?.createdAt || now,
    updatedAt: now,
  };
  const stored = await savePersonel(p);
  if (!stored) {
    return NextResponse.json(
      { ok: false, error: "Kalıcı depolama yapılandırılmadı." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, personel: p });
}
