// /catalogs/*.pdf dosyaları (toptan fiyat listeleri!) public klasöründe durur
// ama giriş yapmadan indirilememeli. Middleware oturum çerezini doğrular;
// geçersizse giriş sayfasına yönlendirir (giriş sonrası aynı dosyaya döner).

import { NextRequest, NextResponse } from "next/server";
import { verifySessionToken, SESSION_COOKIE } from "@/lib/jwt";

export async function middleware(req: NextRequest) {
  const token = req.cookies.get(SESSION_COOKIE)?.value;
  const user = token ? await verifySessionToken(token) : null;
  if (!user) {
    const url = req.nextUrl.clone();
    url.pathname = "/giris";
    url.search = `?next=${encodeURIComponent(req.nextUrl.pathname)}`;
    return NextResponse.redirect(url);
  }
  return NextResponse.next();
}

export const config = {
  matcher: "/catalogs/:path*",
};
