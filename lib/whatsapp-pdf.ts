// Sipariş fişi PDF'ini WhatsApp Cloud API ile patrona/yöneticilere DOSYA olarak
// gönderir. İki adım: (1) PDF, işletme numarasının medya deposuna yüklenir,
// (2) Meta'nın onayladığı şablonla (belge başlıklı) alıcıya gönderilir.
//
// Neden şablon? WhatsApp, işletmenin kendiliğinden yazmasına yalnızca onaylı
// şablonla izin verir; alıcı son 24 saatte yazmadıysa serbest mesaj reddedilir.
// Şablon adı tanımlı değilse serbest belge mesajı denenir (yalnızca 24 saat
// penceresi açıkken çalışır) — kurulum/test aşaması için.
//
// Ortam değişkenleri:
//   WHATSAPP_TOKEN, WHATSAPP_PHONE_ID   — mevcut Cloud API bağlantısı
//   PATRON_WHATSAPP                     — alıcı numara(lar), virgülle: 05325099442,0532...
//   WHATSAPP_TEMPLATE_SIPARIS           — onaylı şablon adı (örn. siparis_fisi)
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

export function sablonAdi(): string {
  return (process.env.WHATSAPP_TEMPLATE_SIPARIS || "").trim();
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

async function grafHata(res: Response): Promise<string> {
  try {
    const j = (await res.json()) as { error?: { message?: string; code?: number; error_data?: { details?: string } } };
    const e = j?.error;
    if (!e) return `HTTP ${res.status}`;
    const detay = e.error_data?.details ? ` — ${e.error_data.details}` : "";
    return `${e.message || "hata"} (kod ${e.code ?? "?"})${detay}`;
  } catch {
    return `HTTP ${res.status}`;
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
  if (!res.ok) return { hata: await grafHata(res) };
  const j = (await res.json()) as { id?: string };
  return j.id ? { id: j.id } : { hata: "Medya kimliği dönmedi." };
}

async function mesajGonder(body: Record<string, unknown>): Promise<{ ok: boolean; hata?: string }> {
  const token = process.env.WHATSAPP_TOKEN!;
  const phoneId = process.env.WHATSAPP_PHONE_ID!;
  const res = await fetch(`${GRAPH}/${phoneId}/messages`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ messaging_product: "whatsapp", recipient_type: "individual", ...body }),
  });
  if (res.ok) return { ok: true };
  return { ok: false, hata: await grafHata(res) };
}

/**
 * Fiş PDF'ini bütün patron numaralarına gönderir. Sipariş kaydını asla
 * engellemez: hata olursa sonuç nesnesinde döner, çağıran loglar.
 */
export async function sendPdfToPatron(pdf: Buffer, bilgi: FisBildirim): Promise<PatronGonderim> {
  const alicilar = patronNumaralari();
  if (!whatsappConfigured() || !alicilar.length) {
    return { ok: false, gonderilen: [], hatalar: ["WhatsApp Cloud API veya PATRON_WHATSAPP tanımlı değil."], yontem: "yok" };
  }
  const dosya = `${bilgi.orderId}.pdf`;
  const up = await uploadWhatsappPdf(pdf, dosya);
  if (!up.id) {
    return { ok: false, gonderilen: [], hatalar: [`PDF yüklenemedi: ${up.hata}`], yontem: "yok" };
  }

  const sablon = sablonAdi();
  const turEtiket = bilgi.tur === "toptan" ? "Toptan" : "Perakende";
  const ozet = `${turEtiket} sipariş ${bilgi.orderId} · ${param(bilgi.musteri, 60)} · ₺${fmtTL(bilgi.tutar)} · ${param(bilgi.calisan, 40)}`;

  const gonderilen: string[] = [];
  const hatalar: string[] = [];
  for (const to of alicilar) {
    const r = sablon
      ? await mesajGonder({
          to,
          type: "template",
          template: {
            name: sablon,
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
        })
      : await mesajGonder({
          to,
          type: "document",
          document: { id: up.id, filename: dosya, caption: ozet },
        });
    if (r.ok) gonderilen.push(to);
    else hatalar.push(`${to}: ${r.hata}`);
  }
  return { ok: gonderilen.length > 0, gonderilen, hatalar, yontem: sablon ? "sablon" : "serbest" };
}
