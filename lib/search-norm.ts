// Fiyat listesi ve ürün aramalarında ortak metin normalizasyonu.
//
// Depoda telefondan arayan biri Türkçe klavyeyle uğraşmasın, boşluk/tire koymayı
// hatırlamak zorunda kalmasın diye arama iki aşamalı:
//   norm()     — Türkçe karakterleri ASCII'ye çevirir, küçük harfe indirir
//   sikistir() — buna ek olarak harf ve rakam dışındaki her şeyi atar
//
// Böylece "2315S" → "2315 S", "cass7kart" → "CASS-7-KART", "askıteli" → "Askı Teli"
// hepsi bulunur.

const TR_TO_ASCII: Record<string, string> = {
  ç: "c", Ç: "c", ğ: "g", Ğ: "g", ı: "i", İ: "i",
  ö: "o", Ö: "o", ş: "s", Ş: "s", ü: "u", Ü: "u",
};

export function norm(s: string): string {
  return String(s || "")
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_TO_ASCII[c])
    .toLowerCase()
    .trim();
}

export function sikistir(s: string): string {
  return norm(s).replace(/[^a-z0-9]/g, "");
}

// Aranan metin, hedef alanlardan herhangi birinde geçiyor mu?
// Hem normal hem sıkıştırılmış biçimde karşılaştırır.
export function eslesir(sorgu: string, ...alanlar: (string | undefined)[]): boolean {
  const q = norm(sorgu);
  if (!q) return true;
  const qs = sikistir(sorgu);
  return alanlar.some((a) => {
    if (!a) return false;
    return norm(a).includes(q) || (qs.length > 0 && sikistir(a).includes(qs));
  });
}
