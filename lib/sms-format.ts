// SMS metin ve numara yardımcıları — saf fonksiyonlar, hem sunucuda hem
// tarayıcıda kullanılır. Gizli bilgiye dokunan gönderim kodu lib/sms.ts'tedir;
// bu ayrım sayesinde NETGSM kimlik bilgileri istemci paketine hiç girmez.

/**
 * Türkiye cep numarasını NETGSM'in beklediği 10 haneli biçime çevirir: 5XXXXXXXXX
 * "+90 532 123 45 67", "0532 123 45 67", "532-123-4567" → "5321234567"
 * Geçersizse null döner.
 */
export function normalizePhone(raw: string): string | null {
  let d = String(raw || "").replace(/\D/g, "");
  if (d.startsWith("90") && d.length === 12) d = d.slice(2);
  else if (d.startsWith("0") && d.length === 11) d = d.slice(1);
  if (d.length !== 10 || !d.startsWith("5")) return null;
  return d;
}

// GSM-7 alfabesinde bulunmayan Türkçe harfler mesajı UCS-2'ye düşürür ve segment
// başına düşen karakter 160'tan 70'e iner — yani aynı mesaj iki katı krediye
// mal olur. ç, ö, ü GSM-7'de yer aldığı için sayıma dahil değildir.
const TR_ONLY = /[ğĞıİşŞ]/;

const TR_TO_ASCII: Record<string, string> = {
  ç: "c", Ç: "C", ğ: "g", Ğ: "G", ı: "i", İ: "I",
  ö: "o", Ö: "O", ş: "s", Ş: "S", ü: "u", Ü: "U",
};

/** Türkçe harfleri ASCII karşılıklarına çevirir — mesajı ucuzlatmak için. */
export function stripTurkish(s: string): string {
  return String(s || "").replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_TO_ASCII[c]);
}

export interface SmsCount {
  encoding: "TR" | "ASCII";
  chars: number;
  segments: number;
  /** Tek segmentte kalan karakter hakkı — kullanıcıya gösterilir. */
  limit: number;
}

/**
 * Mesajın kaç SMS kredisi harcayacağını hesaplar.
 * Türkçe harf varsa UCS-2 (70 / zincirde 67), yoksa GSM-7 (160 / zincirde 153).
 */
export function smsSegments(text: string): SmsCount {
  const s = String(text || "");
  const chars = s.length;
  const turkish = TR_ONLY.test(s);
  const tek = turkish ? 70 : 160;
  const zincir = turkish ? 67 : 153;
  const segments = chars === 0 ? 0 : chars <= tek ? 1 : Math.ceil(chars / zincir);
  return { encoding: turkish ? "TR" : "ASCII", chars, segments, limit: tek };
}
