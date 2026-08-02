import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { getCustomer, customerTitle } from "@/lib/customers";
import { generateLabelPdf } from "@/lib/label-pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const TR_ASCII: Record<string, string> = {
  ç: "c", Ç: "C", ğ: "g", Ğ: "G", ı: "i", İ: "I",
  ö: "o", Ö: "O", ş: "s", Ş: "S", ü: "u", Ü: "U",
};

// HTTP başlığı yalnızca ASCII taşır — Türkçe adlar dosya adında bozulmasın.
function asciiFileName(s: string): string {
  return s
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_ASCII[c] || c)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Za-z0-9 _-]/g, "")
    .trim()
    .replace(/\s+/g, "_");
}

// Kargo etiketi PDF'i: /api/etiket/pdf?id=C123&adet=2
export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^[A-Za-z0-9]{2,40}$/.test(id)) {
    return NextResponse.json({ error: "Geçersiz müşteri" }, { status: 400 });
  }
  const count = Math.max(1, Math.min(50, Number(req.nextUrl.searchParams.get("adet")) || 1));

  const customer = await getCustomer(id);
  if (!customer) {
    return NextResponse.json({ error: "Müşteri bulunamadı" }, { status: 404 });
  }

  let pdf: Buffer;
  try {
    pdf = await generateLabelPdf(customer, count);
  } catch (e) {
    console.error("Etiket PDF üretilemedi:", e);
    return NextResponse.json(
      { error: "Etiket üretilemedi: " + (e instanceof Error ? e.message : "bilinmeyen hata") },
      { status: 500 }
    );
  }

  const title = customerTitle(customer);
  const ascii = asciiFileName(title) || customer.id;
  const raw = `etiket_${title}.pdf`;
  return new NextResponse(new Uint8Array(pdf), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition":
        `attachment; filename="etiket_${ascii}.pdf"; ` +
        `filename*=UTF-8''${encodeURIComponent(raw)}`,
    },
  });
}
