import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { parseStockWorkbook, type StockData } from "@/lib/stock-parse";
import snapshot from "@/data/stock-snapshot.json";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const BLOB_PATH = "stock/latest.json";

// Vercel Blob erişimi iki şekilde çalışır:
// 1) OIDC: depo projeye "Connect" edilmişse (BLOB_STORE_ID otomatik tanımlanır)
// 2) Klasik token: BLOB_READ_WRITE_TOKEN
function blobConfigured(): boolean {
  return Boolean(process.env.BLOB_STORE_ID || process.env.BLOB_READ_WRITE_TOKEN);
}

async function readFromBlob(): Promise<StockData | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const result = await get(BLOB_PATH, { access: "private", useCache: false });
    if (!result || result.statusCode !== 200 || !result.stream) return null;
    const text = await new Response(result.stream).text();
    return JSON.parse(text) as StockData;
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

  if (!blobConfigured()) {
    return NextResponse.json(
      {
        ok: false,
        error:
          "Kalıcı depolama bağlı değil. Vercel'de Storage → Blob deposu → Connect Project adımını yapın, sonra tekrar yükleyin.",
        parsedCount: data.items.length,
      },
      { status: 503 }
    );
  }

  try {
    const { put } = await import("@vercel/blob");
    await put(BLOB_PATH, JSON.stringify(data), {
      access: "private",
      contentType: "application/json",
      addRandomSuffix: false,
      allowOverwrite: true,
    });
  } catch (err) {
    console.error("Blob yazılamadı:", err);
    return NextResponse.json(
      {
        ok: false,
        error:
          "Stok kaydedilemedi. Depo bağlantısını kontrol edin (Storage → Connect Project) ve yeniden deploy edin.",
      },
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
