// Parola özetleme (scrypt) — USERS_JSON'da düz "password" yerine
// "passwordHash" alanı kullanılabilsin diye. Düz parola desteği geriye dönük
// uyumluluk için korunur; ikisi bir arada da çalışır (hash öncelikli).
//
// Hash üretmek için: node scripts/parola-hash.js <parola>
// Çıkan değeri USERS_JSON'daki kullanıcıya "passwordHash" olarak ekleyin ve
// "password" alanını silin.

import { scryptSync, timingSafeEqual, randomBytes } from "crypto";

const KEYLEN = 64;

export function hashPassword(plain: string): string {
  const salt = randomBytes(16).toString("hex");
  const hash = scryptSync(plain, salt, KEYLEN).toString("hex");
  return `scrypt:${salt}:${hash}`;
}

export function verifyPassword(
  stored: { password?: string; passwordHash?: string },
  plain: string
): boolean {
  if (!plain) return false;
  const h = stored.passwordHash;
  if (h && h.startsWith("scrypt:")) {
    const [, salt, hex] = h.split(":");
    if (!salt || !hex) return false;
    try {
      const calc = scryptSync(plain, salt, hex.length / 2);
      return timingSafeEqual(calc, Buffer.from(hex, "hex"));
    } catch {
      return false;
    }
  }
  return (
    typeof stored.password === "string" &&
    stored.password.length > 0 &&
    stored.password === plain
  );
}
