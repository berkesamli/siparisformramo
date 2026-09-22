import { NextRequest, NextResponse } from "next/server";
import { taslakUret } from "@/lib/mesaj/taslak";
import { hataYaniti, mesajKullanici } from "@/lib/mesaj/yetki";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Yapay zekâ yanıt taslağı — yalnızca metin döner; gönderim çalışanın elindedir.
export async function POST(req: NextRequest) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as { konusmaId?: string } | null;
  if (!b?.konusmaId) return NextResponse.json({ ok: false, error: "konusmaId gerekli." }, { status: 400 });
  try {
    const r = await taslakUret(String(b.konusmaId));
    return NextResponse.json({ ok: true, taslak: r.taslak });
  } catch (e) {
    console.error("Taslak üretilemedi:", e);
    return hataYaniti(e, 502);
  }
}
