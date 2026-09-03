import { NextResponse } from "next/server";
import { getDealerSession } from "@/lib/auth";
import {
  getDealerPricing,
  saveDealerPricing,
  saveDealer,
  publicDealer,
  hashPassword,
  verifyPassword,
} from "@/lib/dealers";
import { getUsdRate } from "@/lib/kur";
import { blobConfigured } from "@/lib/store";

export const dynamic = "force-dynamic";

const s = (v: unknown, max = 200) => String(v ?? "").trim().slice(0, max);

export async function GET() {
  const sess = await getDealerSession();
  if (!sess) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  const [pricing, kur] = await Promise.all([getDealerPricing(sess.dealer.slug), getUsdRate()]);
  return NextResponse.json({
    ok: true,
    dealer: publicDealer(sess.dealer),
    pricing,
    autoRate: kur?.usd ?? null,
    blob: blobConfigured(),
  });
}

// Fiyat ayarları + firma bilgileri + şifre değişikliği (tek uç, alanlar opsiyonel)
export async function PUT(req: Request) {
  const sess = await getDealerSession();
  if (!sess) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  if (!blobConfigured()) {
    return NextResponse.json({ error: "Kalıcı depolama (Blob) yapılandırılmamış." }, { status: 503 });
  }
  const body = await req.json().catch(() => null);
  if (!body) return NextResponse.json({ error: "Geçersiz istek" }, { status: 400 });

  const d = sess.dealer;
  let pricing = null;
  if (body.pricing) pricing = await saveDealerPricing(d.slug, body.pricing);

  if (body.profile) {
    const p = body.profile;
    if (s(p.name, 80)) d.name = s(p.name, 80);
    if (p.phone !== undefined) d.phone = s(p.phone, 40);
    if (p.email !== undefined) d.email = s(p.email, 120);
    if (p.address !== undefined) d.address = s(p.address, 240);
    if (p.city !== undefined) d.city = s(p.city, 60);
    if (p.website !== undefined) d.website = s(p.website, 120);
    if (p.contactName !== undefined) d.contactName = s(p.contactName, 80);
  }

  if (body.newPassword) {
    const cur = String(body.currentPassword || "");
    if (!verifyPassword(cur, d.passwordHash)) {
      return NextResponse.json({ error: "Mevcut şifre hatalı." }, { status: 400 });
    }
    const np = String(body.newPassword);
    if (np.length < 6) {
      return NextResponse.json({ error: "Yeni şifre en az 6 karakter olmalı." }, { status: 400 });
    }
    d.passwordHash = hashPassword(np);
  }

  if (body.profile || body.newPassword) await saveDealer(d);

  return NextResponse.json({ ok: true, dealer: publicDealer(d), pricing });
}
