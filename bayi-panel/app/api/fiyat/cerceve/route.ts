import { NextRequest, NextResponse } from "next/server";
import { getDealerSession } from "@/lib/auth";
import { getDealerPricing } from "@/lib/dealers";
import { dealerFramePrice } from "@/lib/frame-price";

export const dynamic = "force-dynamic";

// Seri kodu → bayinin çerçeve metre fiyatı (çarpan sunucuda uygulanır).
export async function GET(req: NextRequest) {
  const s = await getDealerSession();
  if (!s) return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });

  const code = (req.nextUrl.searchParams.get("code") || "").trim();
  const rate = Number(req.nextUrl.searchParams.get("rate")) || 0;
  if (!code || !(rate > 0)) return NextResponse.json({ found: false, tlPerM: 0, costTlPerM: 0 });

  const pricing = await getDealerPricing(s.dealer.slug);
  return NextResponse.json(dealerFramePrice(code, rate, pricing.frameFactor));
}
