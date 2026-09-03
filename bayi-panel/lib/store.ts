// Kalıcı depolama — Vercel Blob üzerinde JSON dosyaları.
// İleride Postgres'e geçilecekse yalnızca bu dosya değişir.

export function blobConfigured(): boolean {
  return Boolean(process.env.BLOB_STORE_ID || process.env.BLOB_READ_WRITE_TOKEN);
}

export async function readJson<T>(path: string): Promise<T | null> {
  if (!blobConfigured()) return null;
  try {
    const { get } = await import("@vercel/blob");
    const r = await get(path, { access: "private", useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return JSON.parse(await new Response(r.stream).text()) as T;
  } catch {
    return null;
  }
}

export async function writeJson(path: string, data: unknown): Promise<boolean> {
  if (!blobConfigured()) return false;
  const { put } = await import("@vercel/blob");
  await put(path, JSON.stringify(data), {
    access: "private",
    contentType: "application/json",
    addRandomSuffix: false,
    allowOverwrite: true,
  });
  return true;
}

export async function deleteJson(path: string): Promise<void> {
  if (!blobConfigured()) return;
  try {
    const { del } = await import("@vercel/blob");
    await del(path);
  } catch {
    /* yoksa geç */
  }
}

export async function listPaths(prefix: string, limit = 1000): Promise<string[]> {
  if (!blobConfigured()) return [];
  try {
    const { list } = await import("@vercel/blob");
    const { blobs } = await list({ prefix, limit });
    return blobs.map((b) => b.pathname);
  } catch {
    return [];
  }
}

export async function readMany<T>(paths: string[]): Promise<T[]> {
  const out: T[] = [];
  await Promise.all(
    paths.map(async (p) => {
      const v = await readJson<T>(p);
      if (v) out.push(v);
    })
  );
  return out;
}

export function istanbulDateKey(d = new Date()): string {
  return d.toLocaleDateString("en-CA", { timeZone: "Europe/Istanbul" });
}

/** Bugünden geriye n günlük tarih anahtarları (İstanbul saati). */
export function lastNDateKeys(n: number): string[] {
  const keys: string[] = [];
  const now = Date.now();
  for (let i = 0; i < n; i++) keys.push(istanbulDateKey(new Date(now - i * 86400000)));
  return keys;
}
