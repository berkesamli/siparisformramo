// Süreç içi kısa ömürlü önbellek — Blob okumalarını (indeks, müşteri listesi,
// stok) aynı sunucu örneğinde birkaç saniye paylaşır. Vercel'de sıcak
// örnekler arasında korunur; soğuk başlangıçta boş başlar (zararsız).

const store = new Map<string, { ts: number; value: unknown; pending?: Promise<unknown> }>();

export async function memo<T>(key: string, ttlMs: number, fn: () => Promise<T>): Promise<T> {
  const now = Date.now();
  const hit = store.get(key);
  if (hit && now - hit.ts < ttlMs && hit.value !== undefined) return hit.value as T;
  if (hit?.pending) return hit.pending as Promise<T>;
  const pending = fn()
    .then((v) => {
      store.set(key, { ts: Date.now(), value: v });
      return v;
    })
    .catch((e) => {
      store.delete(key);
      throw e;
    });
  store.set(key, { ts: hit?.ts || 0, value: hit?.value, pending });
  return pending;
}

export function bust(prefix: string) {
  for (const k of [...store.keys()]) if (k.startsWith(prefix)) store.delete(k);
}
