import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner, isUretimci } from "@/data/users";
import { hataYaniti } from "@/lib/uretim/yetki";
import { dbConfigured } from "@/lib/uretim/db";
import { ikasBen, ikasConfigured, ikasSiparisListesi, ikasWebhookKey, ikasWebhookKur, ikasWebhookListesi, ikasWebhookSil, ikasZaman } from "@/lib/ikas/client";
import { ikasSiparisKalemleri } from "@/lib/ikas/not";
import { ikasSiparisNoIsle } from "@/lib/ikas/senk";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// ikas bağlantı sınaması (sahipler + üretim yetkilileri): jeton, son siparişler ve
// not çözümleme önizlemesi; webhook kurulumu. Ayarlar sayfasındaki kart buradan okur.
async function yetkili() {
  const u = await getSessionUser();
  return u && u.role === "staff" && (isOwner(u.username) || isUretimci(u.username)) ? u : null;
}

function siteAdresi(req: NextRequest): string {
  const env = (process.env.SITE_URL || process.env.NEXT_PUBLIC_SITE_URL || "").trim().replace(/\/$/, "");
  if (env) return env;
  const host = req.headers.get("x-forwarded-host") || req.headers.get("host") || "";
  const proto = req.headers.get("x-forwarded-proto") || "https";
  return host ? `${proto}://${host}` : "";
}

export async function GET(req: NextRequest) {
  const u = await yetkili();
  if (!u) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  const durum = { ikas: ikasConfigured(), webhookKey: Boolean(ikasWebhookKey()), db: dbConfigured(), webhookAdresi: siteAdresi(req) ? `${siteAdresi(req)}/api/ikas/webhook?k=${ikasWebhookKey() ? "•••" : "<IKAS_WEBHOOK_KEY>"}` : "" };
  if (!ikasConfigured()) return NextResponse.json({ ok: false, durum, error: "IKAS_CLIENT_ID / IKAS_CLIENT_SECRET tanımlı değil." });
  try {
    const me = await ikasBen();
    const liste = await ikasSiparisListesi({ limit: 3 });
    const siparisler = liste.data.map((o) => {
      const c = ikasSiparisKalemleri(o);
      return {
        id: o.id, no: o.orderNumber, tarih: ikasZaman(o.orderedAt), durum: o.status, paket: o.orderPackageStatus, kanal: o.salesChannel?.name || o.salesChannelId,
        musteri: o.customer?.fullName || [o.shippingAddress?.firstName, o.shippingAddress?.lastName].filter(Boolean).join(" "),
        tutar: o.totalFinalPrice, satirlar: (o.orderLineItems || []).map((s) => ({ ad: s.variant?.name, sku: s.variant?.sku, adet: s.quantity, secenekler: s.options })),
        not: o.note, oznitelikler: o.attributes, cozum: { kalemler: c.kalemler.map((k) => ({ ...k, retail: k.retail ? "var" : undefined })), eksik: c.eksik, notVar: Boolean(c.notMetni) },
      };
    });
    let webhooklar: unknown = null;
    try { webhooklar = await ikasWebhookListesi(); } catch (e) { webhooklar = { hata: (e as Error)?.message }; }
    return NextResponse.json({ ok: true, durum, me, toplamSiparis: liste.count, siparisler, webhooklar });
  } catch (e) {
    return NextResponse.json({ ok: false, durum, error: (e as Error)?.message || String(e) });
  }
}

// POST { islem: "webhook-kur" | "webhook-sil" | "siparis-al", orderNumber? }
export async function POST(req: NextRequest) {
  const u = await yetkili();
  if (!u) return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  if (!ikasConfigured()) return NextResponse.json({ ok: false, error: "ikas bağlantısı ayarlanmamış." }, { status: 503 });
  const b = (await req.json().catch(() => null)) as { islem?: string; orderNumber?: string; endpoint?: string } | null;
  try {
    if (b?.islem === "webhook-kur") {
      if (!ikasWebhookKey()) return NextResponse.json({ ok: false, error: "Önce IKAS_WEBHOOK_KEY tanımlayın." }, { status: 400 });
      const site = String(b.endpoint || siteAdresi(req)).replace(/\/$/, "");
      if (!/^https:\/\//.test(site)) return NextResponse.json({ ok: false, error: "Webhook adresi https olmalı (SITE_URL)." }, { status: 400 });
      const endpoint = `${site}/api/ikas/webhook?k=${encodeURIComponent(ikasWebhookKey())}`;
      const r = await ikasWebhookKur(endpoint);
      return NextResponse.json({ ok: true, webhooklar: r });
    }
    if (b?.islem === "webhook-sil") return NextResponse.json({ ok: await ikasWebhookSil() });
    if (b?.islem === "siparis-al") {
      if (!dbConfigured()) return NextResponse.json({ ok: false, error: "DATABASE_URL yok." }, { status: 503 });
      const r = await ikasSiparisNoIsle(String(b.orderNumber || ""), u.name);
      return r ? NextResponse.json({ ok: true, sonuc: r }) : NextResponse.json({ ok: false, error: "Sipariş bulunamadı." }, { status: 404 });
    }
    return NextResponse.json({ ok: false, error: "Geçersiz işlem." }, { status: 400 });
  } catch (e) {
    return hataYaniti(e);
  }
}
