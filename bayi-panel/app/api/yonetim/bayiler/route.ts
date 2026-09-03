import { NextRequest, NextResponse } from "next/server";
import { getAdminSession } from "@/lib/auth";
import {
  listDealers,
  getDealer,
  saveDealer,
  deleteDealer,
  rebuildDealerIndex,
  hashPassword,
  slugify,
  normalizeUsername,
  listDealerIndex,
  publicDealer,
  SUBSCRIPTION_LABELS,
  type Dealer,
  type Subscription,
  type SubscriptionStatus,
} from "@/lib/dealers";
import { listAllOrders } from "@/lib/orders";
import { blobConfigured } from "@/lib/store";

export const dynamic = "force-dynamic";

const s = (v: unknown, max = 200) => String(v ?? "").trim().slice(0, max);

function parseSubscription(raw: any, prev?: Subscription): Subscription {
  const status = (raw?.status in SUBSCRIPTION_LABELS ? raw.status : prev?.status || "aktif") as SubscriptionStatus;
  const paidUntil = /^\d{4}-\d{2}-\d{2}$/.test(String(raw?.paidUntil || "")) ? raw.paidUntil : prev?.paidUntil;
  const feeRaw = raw?.monthlyFee;
  const monthlyFee =
    feeRaw === undefined || feeRaw === "" ? prev?.monthlyFee : Math.max(0, Number(feeRaw) || 0);
  const note = raw?.note !== undefined ? s(raw.note, 300) : prev?.note;
  return { status, paidUntil, monthlyFee, note };
}

export async function GET(req: NextRequest) {
  if (!(await getAdminSession())) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  if (req.nextUrl.searchParams.get("rebuild") === "1") await rebuildDealerIndex();

  const dealers = await listDealers();
  const withStats = await Promise.all(
    dealers.map(async (d) => {
      const orders = await listAllOrders(d.slug, 1000);
      const since = Date.now() - 30 * 86400000;
      const last30 = orders.filter((o) => new Date(o.createdAt).getTime() >= since);
      return {
        ...publicDealer(d),
        stats: {
          orders: orders.length,
          orders30: last30.length,
          revenue30: Math.round(last30.reduce((sum, o) => sum + (o.status === "İptal" ? 0 : o.total), 0)),
          lastOrderAt: orders[0]?.createdAt || null,
        },
      };
    })
  );
  return NextResponse.json({
    ok: true,
    dealers: withStats,
    blob: blobConfigured(),
    defaults: {
      fee: Number(process.env.SUBSCRIPTION_FEE_TL) || 0,
      threshold: Number(process.env.FREE_THRESHOLD_TL) || 0,
    },
  });
}

// Yeni bayi
export async function POST(req: Request) {
  if (!(await getAdminSession())) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  if (!blobConfigured()) return NextResponse.json({ error: "Blob yapılandırılmamış" }, { status: 503 });
  const b = await req.json().catch(() => null);
  if (!b) return NextResponse.json({ error: "Geçersiz istek" }, { status: 400 });

  const name = s(b.name, 80);
  const username = normalizeUsername(s(b.username, 40));
  const password = String(b.password || "");
  if (!name || !username || password.length < 6) {
    return NextResponse.json({ error: "Firma adı, kullanıcı adı ve en az 6 karakterli şifre gerekli." }, { status: 400 });
  }
  let slug = slugify(s(b.slug, 40) || name);
  if (slug.length < 2) slug = "bayi-" + Date.now().toString(36);

  const idx = await listDealerIndex();
  if (idx.some((r) => r.slug === slug)) {
    return NextResponse.json({ error: `"${slug}" bayi kodu zaten kullanılıyor.` }, { status: 409 });
  }
  if (idx.some((r) => normalizeUsername(r.username) === username)) {
    return NextResponse.json({ error: "Bu kullanıcı adı zaten kullanılıyor." }, { status: 409 });
  }

  const now = new Date().toISOString();
  const dealer: Dealer = {
    slug,
    username,
    passwordHash: hashPassword(password),
    name,
    contactName: s(b.contactName, 80),
    phone: s(b.phone, 40),
    email: s(b.email, 120),
    address: s(b.address, 240),
    city: s(b.city, 60),
    website: s(b.website, 120),
    active: b.active === undefined ? true : Boolean(b.active),
    subscription: parseSubscription(b.subscription, {
      status: "aktif",
      monthlyFee: Number(process.env.SUBSCRIPTION_FEE_TL) || undefined,
    }),
    createdAt: now,
    updatedAt: now,
  };
  await saveDealer(dealer);
  return NextResponse.json({ ok: true, dealer: publicDealer(dealer) });
}

// Bayi güncelle (bilgiler, durum, abonelik, şifre sıfırlama)
export async function PATCH(req: Request) {
  if (!(await getAdminSession())) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  const b = await req.json().catch(() => null);
  const d = b?.slug ? await getDealer(String(b.slug)) : null;
  if (!d) return NextResponse.json({ error: "Bayi bulunamadı" }, { status: 404 });

  if (b.name !== undefined && s(b.name, 80)) d.name = s(b.name, 80);
  if (b.contactName !== undefined) d.contactName = s(b.contactName, 80);
  if (b.phone !== undefined) d.phone = s(b.phone, 40);
  if (b.email !== undefined) d.email = s(b.email, 120);
  if (b.address !== undefined) d.address = s(b.address, 240);
  if (b.city !== undefined) d.city = s(b.city, 60);
  if (b.website !== undefined) d.website = s(b.website, 120);
  if (b.active !== undefined) d.active = Boolean(b.active);
  if (b.subscription) d.subscription = parseSubscription(b.subscription, d.subscription);
  if (b.username !== undefined) {
    const u = normalizeUsername(s(b.username, 40));
    if (u && u !== d.username) {
      const idx = await listDealerIndex();
      if (idx.some((r) => r.slug !== d.slug && normalizeUsername(r.username) === u)) {
        return NextResponse.json({ error: "Bu kullanıcı adı zaten kullanılıyor." }, { status: 409 });
      }
      d.username = u;
    }
  }
  if (b.newPassword) {
    if (String(b.newPassword).length < 6) {
      return NextResponse.json({ error: "Şifre en az 6 karakter olmalı." }, { status: 400 });
    }
    d.passwordHash = hashPassword(String(b.newPassword));
  }

  await saveDealer(d);
  return NextResponse.json({ ok: true, dealer: publicDealer(d) });
}

export async function DELETE(req: NextRequest) {
  if (!(await getAdminSession())) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  const slug = req.nextUrl.searchParams.get("slug") || "";
  const d = await getDealer(slug);
  if (!d) return NextResponse.json({ error: "Bayi bulunamadı" }, { status: 404 });
  // Siparişler silinmez (arşiv olarak kalır); yalnızca hesap ve indeks kaydı kaldırılır.
  await deleteDealer(slug);
  return NextResponse.json({ ok: true });
}
