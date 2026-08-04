import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isFinance, isOwner } from "@/data/users";
import { getOzetRange, rebuildOzet } from "@/lib/finans-ozet";
import { listCekSenet } from "@/lib/ceksenet";
import { istanbulDateKey } from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Finans genel bakış verisi — yalnızca özet dosyaları + çek portföyü okunur;
// tahsilat/gider kayıtları taranmaz (Blob tutumluluğu).
export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isFinance(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }

  // Son 12 ay (bu ay dahil)
  const bugun = istanbulDateKey();
  const [y, m] = bugun.split("-").map(Number);
  const months: string[] = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(Date.UTC(y, m - 1 - i, 1));
    months.push(
      `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`
    );
  }

  const [ozetler, cekler] = await Promise.all([
    getOzetRange(months),
    listCekSenet(),
  ]);

  const portfoy = cekler.filter((c) => c.durum === "portfoyde");
  const otuzGun = new Date(Date.now() + 30 * 86400000)
    .toISOString()
    .slice(0, 10);
  const vadesiYaklasan = portfoy
    .filter((c) => c.vade <= otuzGun)
    .slice(0, 20)
    .map((c) => ({
      id: c.id,
      tur: c.tur,
      kind: c.kind,
      vade: c.vade,
      tutar: c.tutar,
      kimden: c.customerName || c.supplier || "-",
      branch: c.branch,
      gecmis: c.vade < bugun,
    }));

  const r2 = (n: number) => Math.round(n * 100) / 100;
  const portfoyOzet = {
    alinanAdet: portfoy.filter((c) => c.tur === "alinan").length,
    alinanToplam: r2(
      portfoy.filter((c) => c.tur === "alinan").reduce((s, c) => s + c.tutar, 0)
    ),
    verilenAdet: portfoy.filter((c) => c.tur === "verilen").length,
    verilenToplam: r2(
      portfoy.filter((c) => c.tur === "verilen").reduce((s, c) => s + c.tutar, 0)
    ),
  };

  return NextResponse.json({
    ok: true,
    months,
    ozetler,
    vadesiYaklasan,
    portfoyOzet,
  });
}

// Özet tamiri — yalnızca sahipler.
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as { months?: string[] } | null;
  const months = (body?.months || []).filter((m) => /^\d{4}-\d{2}$/.test(m));
  if (!months.length) {
    return NextResponse.json({ ok: false, error: "months gerekli." }, { status: 400 });
  }
  for (const m of months) await rebuildOzet(m);
  return NextResponse.json({ ok: true, rebuilt: months });
}
