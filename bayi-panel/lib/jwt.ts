// Edge-uyumlu JWT yardımcıları (next/headers içermez).
import { SignJWT, jwtVerify } from "jose";

export type SessionKind = "admin" | "dealer";

export interface SessionUser {
  kind: SessionKind;
  username: string;
  name: string;
  slug?: string; // bayi kodu (kind = dealer)
}

export const SESSION_COOKIE = "olga_bayi_session";

export function authSecret(): string {
  return process.env.AUTH_SECRET || "olga-bayi-dev-secret-change-me";
}

function key(): Uint8Array {
  return new TextEncoder().encode(authSecret());
}

export async function createSessionToken(user: SessionUser): Promise<string> {
  return new SignJWT({ ...user })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime("14d")
    .sign(key());
}

export async function verifySessionToken(token: string): Promise<SessionUser | null> {
  try {
    const { payload } = await jwtVerify(token, key());
    if (typeof payload.username !== "string") return null;
    const kind = payload.kind === "admin" ? "admin" : "dealer";
    return {
      kind,
      username: payload.username,
      name: (payload.name as string) || payload.username,
      slug: typeof payload.slug === "string" ? payload.slug : undefined,
    };
  } catch {
    return null;
  }
}
