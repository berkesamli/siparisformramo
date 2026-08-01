import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { parseStockWorkbook, type StockData } from "@/lib/stock-parse";
import snapshot from "@/data/stock-snapshot.json";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const BLOB_PATH = "stock/latest.json";

async function readFromBlob(): Promise<StockData | null> {
  if (!process.env.BLOB_READ_WRITE_TOKEN) return null;
  try {
    const { list } = await import("@vercel/blob");
    const { blobs } = await list({ prefix: BLOB_PATH, limit: 1 });
    if (!blobs.length) return null;
    const res = await fetch(blobs[0].url, { cache: "no-store" });
    if (!res.ok) return null;
    return (await res.json()) as StockData;
  } catch (err) {
    console.error("Blob okunamadı:", err);
    return null;
  }
}

// Güncel stok verisi: önce Blob (günlük yüklenen), yoksa repo içindeki snapshot.
export async function GET() {
  const user = await getSessionUser();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 });
  }
  const data = (await readFromBlob()) ?? (snapshot as StockData);
  return NextResponse.json({ ok: true, data });
}

// Günlük Excel yükleme — yalnızca çalışanlar.
export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }

  const form = await req.formData().catch(() => null);
  const file = form?.get("file");
  if (!file || typeof file === "string") {
    return NextResponse.json({ ok: false, error: "Dosya seçilmedi." }, { status: 400 });
  }

  let data: StockData;
  try {
    const buffer = Buffer.from(await file.arrayBuffer());
    data = parseStockWorkbook(buffer, file.name);
  } catch (err) {
    return NextResponse.json(
      { ok: false, error: err instanceof Error ? err.message : "Excel okunamadı." },
      { status: 400 }
    );
  }

  if (!process.env.BLOB_READ_WRITE_TOKEN) {
    return NextResponse.json(
      {
        ok: false,
        error:
          "Kalıcı depolama henüz açık değil. Vercel panelinde Storage → Create Database → Blob oluşturup projeye bağlayın (BLOB_READ_WRITE_TOKEN otomatik eklenir), sonra tekrar yükleyin.",
        parsedCount: data.items.length,
      },
      { status: 503 }
    );
  }

  try {
    const { put } = await import("@vercel/blob");
    await put(BLOB_PATH, JSON.stringify(data), {
      access: "public",
      contentType: "application/json",
      addRandomSuffix: false,
      allowOverwrite: true,
    });
  } catch (err) {
    console.error("Blob yazılamadı:", err);
    return NextResponse.json(
      { ok: false, error: "Stok kaydedilemedi (Blob hatası)." },
      { status: 500 }
    );
  }

  const ankaraTotal = data.items.reduce((s, i) => s + i.ankaraMt, 0);
  const istanbulTotal = data.items.reduce((s, i) => s + i.istanbulMt, 0);
  return NextResponse.json({
    ok: true,
    count: data.items.length,
    ankaraTotal: Math.round(ankaraTotal),
    istanbulTotal: Math.round(istanbulTotal),
    updatedAt: data.updatedAt,
  });
}
