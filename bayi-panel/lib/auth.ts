// Sunucu tarafı oturum yardımcıları (cookies() kullanır).
import { cookies } from "next/headers";
import { verifySessionToken, SESSION_COOKIE, type SessionUser } from "@/lib/jwt";
import { getDealer, dealerCanLogin, type Dealer } from "@/lib/dealers";

export { createSessionToken, verifySessionToken, SESSION_COOKIE } from "@/lib/jwt";
export type { SessionUser } from "@/lib/jwt";

export async function getSessionUser(): Promise<SessionUser | null> {
  const token = cookies().get(SESSION_COOKIE)?.value;
  if (!token) return null;
  return verifySessionToken(token);
}

export interface DealerSession {
  user: SessionUser;
  dealer: Dealer;
}

/**
 * Bayi oturumu + güncel bayi kaydı. Kayıt silinmiş ya da pasife alınmışsa
 * oturum geçersiz sayılır (çerez süresi dolmamış olsa bile).
 */
export async function getDealerSession(): Promise<DealerSession | null> {
  const user = await getSessionUser();
  if (!user || user.kind !== "dealer" || !user.slug) return null;
  const dealer = await getDealer(user.slug);
  if (!dealer || !dealerCanLogin(dealer)) return null;
  return { user, dealer };
}

export async function getAdminSession(): Promise<SessionUser | null> {
  const user = await getSessionUser();
  return user?.kind === "admin" ? user : null;
}
