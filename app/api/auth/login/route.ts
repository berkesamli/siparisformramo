import { NextResponse } from "next/server";
import { findUser } from "@/data/users";
import { createSessionToken, SESSION_COOKIE } from "@/lib/auth";

export async function POST(req: Request) {
  const body = await req.json().catch(() => null);
  const username = String(body?.username || "").trim();
  const password = String(body?.password || "");

  if (!username || !password) {
    return NextResponse.json({ ok: false, error: "Kullanıcı adı ve şifre gerekli." }, { status: 400 });
  }

  const user = findUser(username, password);
  if (!user) {
    return NextResponse.json({ ok: false, error: "Kullanıcı adı veya şifre hatalı." }, { status: 401 });
  }

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
