import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { retailFramePrice } from "@/lib/retail-orders";

export const dynamic = "force-dynamic";

// Perakende çerçeve metre fiyatı — katsayı sunucuda kalır, istemciye
// yalnızca hesaplanmış TL/m fiyatı döner.
export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const code = (req.nextUrl.searchParams.get("code") || "").trim();
  const rate = Number(req.nextUrl.searchParams.get("rate")) || 0;
  if (!code || !(rate > 0)) {
    return NextResponse.json({ found: false, tlPerM: 0 });
  }

  return NextResponse.json(retailFramePrice(code, rate));
}
