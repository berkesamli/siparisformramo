import { NextResponse } from "next/server";
import type { NextRequest } from "next/server";
import { jwtVerify } from "jose";

// Vercel Edge ortamı middleware içinde "@/..." import'larını çözemediği için
// JWT doğrulama burada kendi içine kapalı tutulur. Çerez adı ve secret
// lib/jwt.ts ile aynı olmalıdır.
const SESSION_COOKIE = "olga_session";

async function verifyToken(
  token: string
): Promise<{ role: string } | null> {
  try {
    const secret = new TextEncoder().encode(
      process.env.AUTH_SECRET || "olga-cerceve-dev-secret-change-me"
    );
    const { payload } = await jwtVerify(token, secret);
    if (typeof payload.username !== "string" || typeof payload.role !== "string")
      return null;
    return { role: payload.role };
  } catch {
    return null;
  }
}

// /panel yalnızca çalışanlara, /portal çalışan + müşteriye açık.
export async function middleware(req: NextRequest) {
  const { pathname } = req.nextUrl;
  const token = req.cookies.get(SESSION_COOKIE)?.value;
  const user = token ? await verifyToken(token) : null;

  if (!user) {
    const url = req.nextUrl.clone();
    url.pathname = "/giris";
    url.searchParams.set("next", pathname);
    return NextResponse.redirect(url);
  }

  if (pathname.startsWith("/panel") && user.role !== "staff") {
    const url = req.nextUrl.clone();
    url.pathname = "/portal";
    url.search = "";
    return NextResponse.redirect(url);
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/panel/:path*", "/panel", "/portal/:path*", "/portal"],
};
