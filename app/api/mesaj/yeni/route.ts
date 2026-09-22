import { NextRequest, NextResponse } from "next/server";
import { yeniWhatsappMesaji } from "@/lib/mesaj/gonder";
import { hataYaniti, mesajKullanici } from "@/lib/mesaj/yetki";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

// Bizim başlattığımız WhatsApp mesajı (müşteri defterinden ya da elle numara).
export async function POST(req: NextRequest) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  const b = (await req.json().catch(() => null)) as { telefon?: string; ad?: string; metin?: string; musteriId?: string | null; musteriTur?: "toptan" | "perakende" | null; taslakAi?: boolean } | null;
  if (!b?.telefon || !b.metin?.trim()) return NextResponse.json({ ok: false, error: "Telefon ve mesaj gerekli." }, { status: 400 });
  try {
    const r = await yeniWhatsappMesaji({
      telefon: String(b.telefon), ad: b.ad ? String(b.ad).slice(0, 120) : undefined, metin: String(b.metin),
      kullanici: { username: y.user.username, name: y.user.name },
      musteriId: b.musteriId || null, musteriTur: b.musteriTur || null, taslakAi: Boolean(b.taslakAi),
    });
    return NextResponse.json({ ok: true, konusma: r.konusma, mesaj: r.mesaj, yontem: r.yontem, sablon: r.sablon });
  } catch (e) {
    console.error("Yeni WhatsApp mesajı gönderilemedi:", e);
    return hataYaniti(e, 502);
  }
}
