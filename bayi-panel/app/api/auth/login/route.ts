import { NextResponse } from "next/server";
import { createSessionToken, SESSION_COOKIE } from "@/lib/jwt";
import { findAdmin } from "@/lib/admins";
import { findDealerByUsername, verifyPassword, dealerCanLogin } from "@/lib/dealers";

export const dynamic = "force-dynamic";

export async function POST(req: Request) {
  const body = await req.json().catch(() => null);
  const username = String(body?.username || "").trim();
  const password = String(body?.password || "");
  if (!username || !password) {
    return NextResponse.json({ ok: false, error: "Kullanıcı adı ve şifre gerekli." }, { status: 400 });
  }

  // 1) Olga yöneticisi
  const admin = findAdmin(username, password);
  let token: string | null = null;
  let kind: "admin" | "dealer" = "dealer";
  let name = "";

  if (admin) {
    kind = "admin";
    name = admin.name;
    token = await createSessionToken({ kind, username: admin.username, name });
  } else {
    // 2) Bayi
    const dealer = await findDealerByUsername(username);
    if (!dealer || !verifyPassword(password, dealer.passwordHash)) {
      return NextResponse.json({ ok: false, error: "Kullanıcı adı veya şifre hatalı." }, { status: 401 });
    }
    if (!dealerCanLogin(dealer)) {
      return NextResponse.json(
        { ok: false, error: "Bayi hesabınız pasif durumda. Lütfen Olga Çerçeve ile iletişime geçin." },
        { status: 403 }
      );
    }
    name = dealer.name;
    token = await createSessionToken({ kind, username: dealer.username, name, slug: dealer.slug });
  }

  const res = NextResponse.json({ ok: true, kind, name });
  res.cookies.set(SESSION_COOKIE, token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    maxAge: 60 * 60 * 24 * 14,
    path: "/",
  });
  return res;
}
