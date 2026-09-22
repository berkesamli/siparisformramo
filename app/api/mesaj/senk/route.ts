import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isMesajci } from "@/data/users";
import { dbConfigured } from "@/lib/mesaj/db";
import { gmailConfigured, gmailSenk } from "@/lib/mesaj/gmail";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Gmail senkronu: cron (Authorization: Bearer CRON_SECRET) ya da yetkili çalışan.
export async function GET(req: NextRequest) {
  const cron = (process.env.CRON_SECRET || "").trim();
  const auth = req.headers.get("authorization") || "";
  let yetkili = Boolean(cron && auth === `Bearer ${cron}`);
  if (!yetkili) {
    const u = await getSessionUser();
    yetkili = Boolean(u && u.role === "staff" && isMesajci(u.username));
  }
  if (!yetkili) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  if (!dbConfigured()) return NextResponse.json({ ok: false, error: "DATABASE_URL yok." }, { status: 503 });
  if (!gmailConfigured()) return NextResponse.json({ ok: true, sonuc: [], not: "GMAIL_HESAPLAR tanımlı değil." });
  const sonuc = await gmailSenk({ zorla: true });
  return NextResponse.json({ ok: true, sonuc });
}
