// Sunucu tarafı oturum yardımcıları (cookies() kullanır — middleware'de KULLANMAYIN;
// middleware lib/jwt.ts'den import eder).

import { cookies } from "next/headers";
import { verifySessionToken, SESSION_COOKIE, type SessionUser } from "@/lib/jwt";
import { userExists } from "@/data/users";

export { createSessionToken, verifySessionToken, SESSION_COOKIE } from "@/lib/jwt";
export type { SessionUser } from "@/lib/jwt";

export async function getSessionUser(): Promise<SessionUser | null> {
  const token = cookies().get(SESSION_COOKIE)?.value;
  if (!token) return null;
  const user = await verifySessionToken(token);
  // Çerez süresi dolmamış olsa da kullanıcı listeden silindiyse oturum
  // geçersizdir — işten ayrılan kişi USERS_JSON'dan çıkarılınca açık
  // oturumları da o anda kapanır.
  if (user && !userExists(user.username)) return null;
  return user;
}
