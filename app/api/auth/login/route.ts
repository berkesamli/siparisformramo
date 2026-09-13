import { NextResponse } from "next/server";
import { findUserByUsername } from "@/data/users";
import { verifyPassword } from "@/lib/password";
import {
  loginLockRemaining,
  registerLoginFail,
  clearLoginFails,
} from "@/lib/login-guard";
import { createSessionToken, SESSION_COOKIE } from "@/lib/auth";

export const runtime = "nodejs";

export async function POST(req: Request) {
  const body = await req.json().catch(() => null);
  const username = String(body?.username || "").trim();
  const password = String(body?.password || "");

  if (!username || !password) {
    return NextResponse.json({ ok: false, error: "Kullanıcı adı ve şifre gerekli." }, { status: 400 });
  }

  // Kaba kuvvet kilidi: 15 dakikada 5 hatalı deneme → 15 dakika bekleme
  const kilitSn = await loginLockRemaining(username);
  if (kilitSn > 0) {
    const dk = Math.ceil(kilitSn / 60);
    return NextResponse.json(
      {
        ok: false,
        error: `Çok fazla hatalı deneme — bu hesap geçici olarak kilitlendi. ${dk} dakika sonra tekrar deneyin.`,
      },
      { status: 429 }
    );
  }

  const user = findUserByUsername(username);
  if (!user || !verifyPassword(user, password)) {
    const kilitlendi = await registerLoginFail(username);
    return NextResponse.json(
      {
        ok: false,
        error: kilitlendi
          ? "Çok fazla hatalı deneme — hesap 15 dakika kilitlendi."
          : "Kullanıcı adı veya şifre hatalı.",
      },
      { status: kilitlendi ? 429 : 401 }
    );
  }

  await clearLoginFails(username);

  const token = await createSessionToken({
    username: user.username,
    name: user.name,
    role: user.role,
  });

  const res = NextResponse.json({ ok: true, role: user.role, name: user.name });
  res.cookies.set(SESSION_COOKIE, token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    maxAge: 60 * 60 * 24 * 7,
    path: "/",
  });
  return res;
}
