// E-posta gövdesi: HTML temizliği (saklamadan önce) ve arayüzde gösterim yardımcıları.
//
// Güvenlik iki katmanlı: burada betik/form/dış kaynak etiketleri ve olay öznitelikleri atılır; arayüz ise
// gövdeyi betiksiz, kum havuzlu (sandbox) bir iframe'de gösterir. Temizlik bir şeyi kaçırsa bile iframe'de
// betik çalışmaz, sayfaya erişemez. Düz metin yedeği için Gmail'in "[image: …]" / takip bağlantısı
// kalıntılarını ayıklayan mailMetinTemizle vardır.

/** Saklanacak HTML üst sınırı (karakter). Daha büyük gövde düz metin olarak kalır. */
export const HTML_AZAMI = 400_000;

// İçeriğiyle birlikte atılan etiketler
const ICERIKLI = /<(script|iframe|object|embed|noscript|template|applet|frame|frameset|video|audio|svg|math|title|button|select|textarea)\b[^>]*>[\s\S]*?<\/\1\s*>/gi;
// Yalnızca etiketi atılan (içeriği kalan) etiketler
const TEK = /<\/?(script|iframe|object|embed|noscript|template|applet|frame|frameset|video|audio|svg|math|title|button|select|textarea|form|link|base|meta|input|html|head|option)\b[^>]*>/gi;
const OLAY = /\s+on[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi;
const HEDEF = /\s+target\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi;
const ADRES = /\s+(href|src|action|formaction|background|poster|xlink:href|srcset|data)\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+))/gi;

/** E-posta HTML'ini saklamak için temizler: betik, form, dış çerçeve, olay öznitelikleri, javascript:/data: adresler. */
export function htmlTemizle(html: string): string {
  let s = String(html || "");
  if (!s.trim()) return "";
  s = s.replace(/<!--[\s\S]*?-->/g, "");
  s = s.replace(/<!DOCTYPE[^>]*>/gi, "");
  s = s.replace(/<\?xml[^>]*\?>/gi, "");
  s = s.replace(ICERIKLI, "");
  s = s.replace(/<body\b([^>]*)>/gi, "<div$1>").replace(/<\/body\s*>/gi, "</div>");
  s = s.replace(TEK, "");
  // Öznitelikler yalnızca etiket içinde temizlenir (metindeki "on…" sözcüklerine dokunulmaz)
  s = s.replace(/<[^>]+>/g, (etiket) => {
    let t = etiket.replace(OLAY, "").replace(HEDEF, "");
    t = t.replace(ADRES, (hepsi, ad: string, _q, a?: string, b?: string, c?: string) => {
      const deger = (a ?? b ?? c ?? "").trim().replace(/[\s\u0000-\u001f]+/g, "").toLowerCase();
      if (/^(javascript|vbscript|livescript|mocha):/.test(deger)) return "";
      if (deger.startsWith("data:") && !(ad.toLowerCase() === "src" && deger.startsWith("data:image/"))) return "";
      return hepsi;
    });
    // style="…expression(…)" / url(javascript:…)
    t = t.replace(/expression\s*\(/gi, "no(").replace(/url\s*\(\s*['"]?\s*(javascript|vbscript):/gi, "url(no:");
    return t;
  });
  // <style> blokları: dış kaynak ve eski IE tuzakları
  s = s.replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi, (blok) => blok.replace(/@import[^;]*;?/gi, "").replace(/expression\s*\(/gi, "no(").replace(/behavior\s*:/gi, "no:").replace(/url\s*\(\s*['"]?\s*(javascript|vbscript):/gi, "url(no:"));
  return s.trim();
}

/**
 * Düz metin yedeği: Gmail'in "[image: …]", "[https://…]" köşeli kalıntıları, gömülü görsel işaretleri ve
 * uzun takip bağlantıları atılır; "<url>" biçimi düz adrese döner; fazla boş satırlar toplanır.
 */
export function mailMetinTemizle(metin: string): string {
  return String(metin || "")
    .replace(/\r\n?/g, "\n")
    .replace(/\[(image|cid):[^\]]*\]/gi, "")
    .replace(/\[https?:\/\/[^\]\s]*\]/gi, "")
    .replace(/<(https?:\/\/[^>\s]*)>/gi, "$1")
    .replace(/https?:\/\/\S{80,}/gi, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/**
 * Arayüzdeki iframe için tam belge: e-postanın kendi stilleri korunur, görseller/tablolar çerçeveye sığar,
 * bağlantılar yeni sekmede açılır (base target). Karanlık temada da beyaz zemin (e-postalar öyle tasarlanır).
 */
export function mailSrcDoc(html: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><base target="_blank"><style>
html,body{margin:0;padding:0;background:#fff;color:#111;color-scheme:light}
body{padding:6px 4px;font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;word-break:break-word;overflow-x:hidden}
img{max-width:100%;height:auto}
table{max-width:100%!important}
a{color:#0b57d0}
</style></head><body>${html}</body></html>`;
}
