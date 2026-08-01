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

function secret(): Uint8Array {
  const s = process.env.AUTH_SECRET || "olga-cerceve-dev-secret-change-me";
  return new TextEncoder().encode(s);
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
