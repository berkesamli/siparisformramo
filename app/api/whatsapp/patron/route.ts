import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import { getOrder, STATUS_LABELS } from "@/lib/orders";
import { getRetailOrder } from "@/lib/retail-orders";
import { generateOrderPdf, pdfOrderFromSaved } from "@/lib/order-pdf";
import { generateRetailPdf } from "@/lib/retail-pdf";
import { patronBildirimHazir, sendPdfToPatron } from "@/lib/whatsapp-pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

const EN_FAZLA = 30;

interface Istek { d: string; id: string }
interface Sonuc { id: string; ok: boolean; yontem?: string; hata?: string }

const gecerli = (x: unknown): x is Istek =>
  !!x && typeof x === "object" &&
  /^\d{4}-\d{2}-\d{2}$/.test(String((x as Istek).d || "")) &&
  /^(OLG|PRK)-[A-Z0-9-]+$/i.test(String((x as Istek).id || ""));

// Kayıtlı sipariş(ler)in fişini patrona WhatsApp ile yeniden gönderir.
// Gövde: { orders: [{ d: "YYYY-MM-DD", id: "OLG-…" | "PRK-…" }, …] } (en fazla 30)
export async function POST(req: Request) {
  const user = await getSessionUser();
  // Yalnızca sahipler (Berke) — patrona gönderim bir yönetim işi
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  if (!patronBildirimHazir()) {
    return NextResponse.json(
      { ok: false, error: "WhatsApp bildirimi kurulu değil: WHATSAPP_TOKEN, WHATSAPP_PHONE_ID ve PATRON_WHATSAPP tanımlı olmalı." },
      { status: 400 }
    );
  }
  const body = (await req.json().catch(() => null)) as { orders?: unknown } | null;
  const istekler = Array.isArray(body?.orders) ? body!.orders.filter(gecerli) : [];
  if (!istekler.length) {
    return NextResponse.json({ ok: false, error: "Gönderilecek sipariş seçilmedi." }, { status: 400 });
  }
  if (istekler.length > EN_FAZLA) {
    return NextResponse.json({ ok: false, error: `Tek seferde en fazla ${EN_FAZLA} fiş gönderilebilir.` }, { status: 400 });
  }

  // Sırayla gönderilir: Meta'nın hız sınırına takılmamak ve sonuçları sipariş
  // sipariş raporlamak için. Bir fişin hatası diğerlerini durdurmaz.
  const sonuclar: Sonuc[] = [];
  for (const { d, id } of istekler) {
    try {
      if (/^OLG-/i.test(id)) {
        const o = await getOrder(d, id);
        if (!o) { sonuclar.push({ id, ok: false, hata: "Sipariş bulunamadı." }); continue; }
        const pdf = await generateOrderPdf(pdfOrderFromSaved(o, STATUS_LABELS[o.status]));
        const r = await sendPdfToPatron(pdf, { tur: "toptan", orderId: o.orderId, musteri: o.customer, tutar: o.net, calisan: o.employee });
        sonuclar.push({ id, ok: r.ok, yontem: r.yontem, hata: r.hatalar.join(" | ") || undefined });
      } else {
        const o = await getRetailOrder(d, id);
        if (!o) { sonuclar.push({ id, ok: false, hata: "Sipariş bulunamadı." }); continue; }
        const pdf = await generateRetailPdf(o);
        const r = await sendPdfToPatron(pdf, { tur: "perakende", orderId: o.orderId, musteri: o.customerName, tutar: o.total, calisan: o.employee });
        sonuclar.push({ id, ok: r.ok, yontem: r.yontem, hata: r.hatalar.join(" | ") || undefined });
      }
    } catch (err) {
      sonuclar.push({ id, ok: false, hata: err instanceof Error ? err.message : "Gönderilemedi." });
    }
  }
  const giden = sonuclar.filter((s) => s.ok).length;
  console.log(`Patrona fiş: ${giden}/${sonuclar.length} gönderildi (${user.name})`);
  return NextResponse.json({ ok: giden > 0, giden, toplam: sonuclar.length, sonuclar });
}
