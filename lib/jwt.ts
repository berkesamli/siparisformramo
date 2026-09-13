// Edge-uyumlu JWT yardımcıları — middleware buradan import eder.
// Bu dosya next/headers gibi yalnızca sunucuda çalışan modüller içermemelidir.

import { SignJWT, jwtVerify } from "jose";
import type { Role } from "@/data/users";

export interface SessionUser {
  username: string;
  name: string;
  role: Role;
}

export const SESSION_COOKIE = "olga_session";

let secretUyarisiYazildi = false;

function secret(): Uint8Array {
  const s = process.env.AUTH_SECRET;
  if (s && s.trim().length >= 16) return new TextEncoder().encode(s.trim());

  // Üretimde AUTH_SECRET tanımlı değilse HERKESÇE BİLİNEN sabit anahtar
  // kullanılmaz: anahtar USERS_JSON içeriğinden türetilir. Bu değer şifreleri
  // içerdiği için dışarıdan bilinemez → oturum çerezi sahte üretilemez.
  // (USERS_JSON değişince herkes yeniden giriş yapar — kabul edilebilir.)
  const users = process.env.USERS_JSON;
  if (process.env.NODE_ENV === "production" && users && users.length >= 32) {
    if (!secretUyarisiYazildi) {
      secretUyarisiYazildi = true;
      console.warn(
        "AUTH_SECRET tanımlı değil — oturum anahtarı USERS_JSON'dan türetiliyor. " +
          "Vercel ortam değişkenlerine kalıcı bir AUTH_SECRET eklemeniz önerilir."
      );
    }
    return new TextEncoder().encode("olga-auth-turetilmis:" + users);
  }

  // Yerel geliştirme / ilk kurulum (yalnızca demo kullanıcı varken)
  return new TextEncoder().encode("olga-cerceve-dev-secret-change-me");
}

export async function createSessionToken(user: SessionUser): Promise<string> {
  return new SignJWT({ username: user.username, name: user.name, role: user.role })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime("7d")
    .sign(secret());
}

export async function verifySessionToken(token: string): Promise<SessionUser | null> {
  try {
    const { payload } = await jwtVerify(token, secret());
    if (typeof payload.username !== "string" || typeof payload.role !== "string") return null;
    return {
      username: payload.username,
      name: (payload.name as string) || payload.username,
      role: payload.role as Role,
    };
  } catch {
    return null;
  }
}
