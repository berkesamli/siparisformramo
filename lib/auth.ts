// Sunucu tarafı oturum yardımcıları (cookies() kullanır — middleware'de KULLANMAYIN;
// middleware lib/jwt.ts'den import eder).

import { cookies } from "next/headers";
import { verifySessionToken, SESSION_COOKIE, type SessionUser } from "@/lib/jwt";

export { createSessionToken, verifySessionToken, SESSION_COOKIE } from "@/lib/jwt";
export type { SessionUser } from "@/lib/jwt";

export async function getSessionUser(): Promise<SessionUser | null> {
  const token = cookies().get(SESSION_COOKIE)?.value;
  if (!token) return null;
  return verifySessionToken(token);
}
