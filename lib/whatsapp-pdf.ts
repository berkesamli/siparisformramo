// Sipariş fişi PDF'ini WhatsApp Cloud API ile patrona/yöneticilere DOSYA olarak
// gönderir. İki adım: (1) PDF, işletme numarasının medya deposuna yüklenir,
// (2) Meta'nın onayladığı şablonla (belge başlıklı) alıcıya gönderilir.
//
// Neden şablon? WhatsApp, işletmenin kendiliğinden yazmasına yalnızca onaylı
// şablonla izin verir; alıcı son 24 saatte yazmadıysa serbest mesaj reddedilir.
// Şablon adı tanımlı değilse ya da Meta şablonu reddederse (henüz onaylanmamış,
// yok, duraklatılmış, parametre uyumsuz) serbest belge mesajı denenir — bu
// yalnızca alıcı son 24 saatte yazdıysa gider; kurulum/onay bekleme aşaması için.
//
// Ortam değişkenleri:
//   WHATSAPP_TOKEN, WHATSAPP_PHONE_ID   — mevcut Cloud API bağlantısı
//   PATRON_WHATSAPP                     — alıcı numara(lar), virgülle: 05325099442,0532...
//   WHATSAPP_TEMPLATE_SIPARIS           — onaylı şablon ad(lar)ı, virgülle: siparis_fisi,siparis_fisi_v2
//                                         (sırayla denenir; ilk kabul edilen kullanılır)
//   WHATSAPP_TEMPLATE_DIL               — şablon dili (varsayılan tr)

const GRAPH = "https://graph.facebook.com/v20.0";

export interface FisBildirim {
  tur: "toptan" | "perakende";
  orderId: string;
  musteri: string;
  tutar: number;   // TL
  calisan: string;
}

export interface PatronGonderim {
  ok: boolean;             // en az bir alıcıya gitti
  gonderilen: string[];    // başarılı numaralar
  hatalar: string[];       // numara: hata
  notlar: string[];        // uyarılar (örn. şablon reddedildi, serbest mesajla gitti)
  yontem: "sablon" | "serbest" | "yok";
}

export function whatsappConfigured(): boolean {
  return Boolean(process.env.WHATSAPP_TOKEN && process.env.WHATSAPP_PHONE_ID);
}

/** "0532 509 94 42" → "905325099442". Boş/geçersizse null. */
export function normalizeWaNumber(raw: string): string | null {
  let d = String(raw || "").replace(/\D/g, "");
  if (!d) return null;
  if (d.startsWith("00")) d = d.slice(2);
  if (d.length === 11 && d.startsWith("0")) d = "90" + d.slice(1);
  else if (d.length === 10 && d.startsWith("5")) d = "90" + d;
  if (d.length < 11 || d.length > 15) return null;
  return d;
}

export function patronNumaralari(): string[] {
  const raw = process.env.PATRON_WHATSAPP || "";
  const out = new Set<string>();
  // Yalnızca virgül/noktalı virgül ayırır; numara içindeki boşluklar normalize edilir
  for (const p of raw.split(/[,;]+/)) {
    const n = normalizeWaNumber(p);
    if (n) out.add(n);
  }
  return [...out];
}

/** Sırayla denenecek şablon adları (virgülle ayrılmış, boşlar atılır). */
export function sablonAdlari(): string[] {
  const out: string[] = [];
  for (const p of (process.env.WHATSAPP_TEMPLATE_SIPARIS || "").split(/[,;]+/)) {
    const ad = p.trim();
    if (ad && !out.includes(ad)) out.push(ad);
  }
  return out;
}
/** Ekranda gösterim için: "siparis_fisi, siparis_fisi_v2" ya da "". */
export function sablonAdi(): string {
  return sablonAdlari().join(", ");
}
export function sablonDili(): string {
  return (process.env.WHATSAPP_TEMPLATE_DIL || "tr").trim();
}

export function patronBildirimHazir(): boolean {
  return whatsappConfigured() && patronNumaralari().length > 0;
}

/** Şablon parametrelerinde satır sonu/sekme yasak, 4+ boşluk yasak, makul uzunluk. */
function param(s: string, max = 80): string {
  const t = String(s ?? "").replace(/[\r\n\t]+/g, " ").replace(/\s{2,}/g, " ").trim();
  return (t || "—").slice(0, max);
}

const fmtTL = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// Sık görülen Meta hata kodları için Türkçe açıklama (ekranda yöneticiye gösterilir).
const HATA_IPUCU: Record<number, string> = {
  190: "token geçersiz ya da süresi dolmuş; Meta'da yeni token üretip WHATSAPP_TOKEN'ı güncelleyin",
  132001: "şablon Meta'da yok ya da henüz onaylanmadı (İnceleniyor); onaylanınca kendiliğinden düzelir",
  132000: "şablondaki değişken sayısı gönderilenle uyuşmuyor; gövdede 4 değişken olmalı",
  132012: "şablon parametre biçimi uyuşmuyor; başlık Belge, gövde 4 metin değişkeni olmalı",
  132015: "şablon Meta tarafından duraklatıldı",
  132016: "şablon Meta tarafından devre dışı bırakıldı",
  131030: "alıcı numara Meta'daki izin listesinde değil (uygulama geliştirme modundayken alıcı listeye eklenmeli)",
  131047: "24 saat penceresi kapalı; alıcı önce işletme numarasına yazmalı ya da onaylı şablon kullanılmalı",
  131026: "mesaj teslim edilemedi; alıcı numara WhatsApp'ta kayıtlı olmayabilir",
};

// Bu kodlarda şablon yerine serbest belge mesajı denemek anlamlı (sorun şablonda).
const SABLON_HATALARI = new Set([132000, 132001, 132005, 132007, 132012, 132015, 132016]);

interface GrafHata { mesaj: string; kod?: number }

async function grafHata(res: Response): Promise<GrafHata> {
  try {
    const j = (await res.json()) as { error?: { message?: string; code?: number; error_data?: { details?: string } } };
    const e = j?.error;
    if (!e) return { mesaj: `HTTP ${res.status}` };
    const detay = e.error_data?.details ? ` — ${e.error_data.details}` : "";
    const ipucu = typeof e.code === "number" && HATA_IPUCU[e.code] ? ` → ${HATA_IPUCU[e.code]}` : "";
    return { mesaj: `${e.message || "hata"} (kod ${e.code ?? "?"})${detay}${ipucu}`, kod: e.code };
  } catch {
    return { mesaj: `HTTP ${res.status}` };
  }
}

/** PDF'i WhatsApp medya deposuna yükler; medya kimliği döner (30 gün geçerli). */
export async function uploadWhatsappPdf(pdf: Buffer, filename: string): Promise<{ id?: string; hata?: string }> {
  const token = process.env.WHATSAPP_TOKEN!;
  const phoneId = process.env.WHATSAPP_PHONE_ID!;
  const form = new FormData();
  form.append("messaging_product", "whatsapp");
  form.append("type", "application/pdf");
  form.append("file", new Blob([new Uint8Array(pdf)], { type: "application/pdf" }), filename);
  const res = await fetch(`${GRAPH}/${phoneId}/media`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body: form,
  });
  if (!res.ok) return { hata: (await grafHata(res)).mesaj };
  const j = (await res.json()) as { id?: string };
  return j.id ? { id: j.id } : { hata: "Medya kimliği dönmedi." };
}

async function mesajGonder(body: Record<string, unknown>): Promise<{ ok: boolean; hata?: string; kod?: number }> {
  const token = process.env.WHATSAPP_TOKEN!;
  const phoneId = process.env.WHATSAPP_PHONE_ID!;
  const res = await fetch(`${GRAPH}/${phoneId}/messages`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ messaging_product: "whatsapp", recipient_type: "individual", ...body }),
  });
  if (res.ok) return { ok: true };
  const h = await grafHata(res);
  return { ok: false, hata: h.mesaj, kod: h.kod };
}

/**
 * Fiş PDF'ini bütün patron numaralarına gönderir. Sipariş kaydını asla
 * engellemez: hata olursa sonuç nesnesinde döner, çağıran loglar.
 */
export async function sendPdfToPatron(pdf: Buffer, bilgi: FisBildirim): Promise<PatronGonderim> {
  const alicilar = patronNumaralari();
  if (!whatsappConfigured() || !alicilar.length) {
    return { ok: false, gonderilen: [], hatalar: ["WhatsApp Cloud API veya PATRON_WHATSAPP tanımlı değil."], notlar: [], yontem: "yok" };
  }
  const dosya = `${bilgi.orderId}.pdf`;
  const up = await uploadWhatsappPdf(pdf, dosya);
  if (!up.id) {
    return { ok: false, gonderilen: [], hatalar: [`PDF yüklenemedi: ${up.hata}`], notlar: [], yontem: "yok" };
  }

  const sablonlar = sablonAdlari();
  const turEtiket = bilgi.tur === "toptan" ? "Toptan" : "Perakende";
  const ozet = `${turEtiket} sipariş ${bilgi.orderId} · ${param(bilgi.musteri, 60)} · ₺${fmtTL(bilgi.tutar)} · ${param(bilgi.calisan, 40)}`;

  const sablonMesaji = (to: string, ad: string) => ({
    to,
    type: "template",
    template: {
      name: ad,
      language: { code: sablonDili() },
      components: [
        { type: "header", parameters: [{ type: "document", document: { id: up.id, filename: dosya } }] },
        {
          type: "body",
          parameters: [
            { type: "text", text: param(`${turEtiket} ${bilgi.orderId}`) },
            { type: "text", text: param(bilgi.musteri, 60) },
            { type: "text", text: fmtTL(bilgi.tutar) },
            { type: "text", text: param(bilgi.calisan, 40) },
          ],
        },
      ],
    },
  });
  const serbestMesaj = (to: string) => ({
    to,
    type: "document",
    document: { id: up.id, filename: dosya, caption: ozet },
  });

  const gonderilen: string[] = [];
  const hatalar: string[] = [];
  const notlar: string[] = [];
  let sablonlaGitti = false;
  let serbestGitti = false;
  for (const to of alicilar) {
    if (!sablonlar.length) {
      const r = await mesajGonder(serbestMesaj(to));
      if (r.ok) { gonderilen.push(to); serbestGitti = true; }
      else hatalar.push(`${to}: ${r.hata}`);
      continue;
    }
    // Şablonlar sırayla: ilk kabul edilen gönderir. Şablon kaynaklı hata
    // (yok/onaysız/uyumsuz) → sıradakine geç; başka hata → dur, raporla.
    const sablonHatalari: string[] = [];
    let gitti = false;
    let sablonSorunu = true;
    for (const ad of sablonlar) {
      const r = await mesajGonder(sablonMesaji(to, ad));
      if (r.ok) {
        gitti = true;
        if (ad !== sablonlar[0]) notlar.push(`${to}: "${ad}" şablonuyla gönderildi (öncekiler kabul edilmedi).`);
        break;
      }
      sablonHatalari.push(`${ad}: ${r.hata}`);
      if (!(r.kod && SABLON_HATALARI.has(r.kod))) { sablonSorunu = false; break; }
    }
    if (gitti) { gonderilen.push(to); sablonlaGitti = true; continue; }
    if (!sablonSorunu) { hatalar.push(`${to}: ${sablonHatalari.join(" · ")}`); continue; }
    // Sorun şablonlarda: 24 saat penceresi açıksa serbest belge gider.
    const r2 = await mesajGonder(serbestMesaj(to));
    if (r2.ok) {
      gonderilen.push(to);
      serbestGitti = true;
      notlar.push(`${to}: şablon kabul edilmedi (${sablonHatalari.join(" · ")}); serbest belge mesajıyla gönderildi.`);
    } else {
      hatalar.push(`${to}: şablon: ${sablonHatalari.join(" · ")} · serbest belge: ${r2.hata}`);
    }
  }
  const yontem: PatronGonderim["yontem"] =
    sablonlaGitti ? "sablon" : serbestGitti ? "serbest" : sablonlar.length ? "sablon" : "serbest";
  return { ok: gonderilen.length > 0, gonderilen, hatalar, notlar, yontem };
}
