// NETGSM SMS gönderimi — YALNIZCA SUNUCU TARAFI.
// Bu dosya NETGSM şifresine dokunur; istemci bileşenlerinden içe aktarmayın.
// Numara/karakter yardımcıları için lib/sms-format.ts kullanın.
//
// Kimlik bilgileri ortam değişkenlerinden okunur (Vercel → Environment Variables):
//   NETGSM_USERCODE  — API alt kullanıcı adı / abone no
//   NETGSM_PASSWORD  — API alt kullanıcı şifresi
//   NETGSM_HEADER    — NETGSM'de onaylanmış gönderici başlığı (ör. OLGACERCEVE)
//
// Not: NETGSM tarafında IP kısıtlaması TANIMLANMAMALIDIR. Vercel'in çıkış IP'si
// sabit değildir; kısıtlama varsa istekler 30 koduyla reddedilir.

import { normalizePhone, smsSegments } from "./sms-format";

const ENDPOINT = "https://api.netgsm.com.tr/sms/rest/v2/send";

export function smsConfigured(): boolean {
  return Boolean(
    process.env.NETGSM_USERCODE &&
      process.env.NETGSM_PASSWORD &&
      process.env.NETGSM_HEADER
  );
}

// NETGSM yanıt kodları. Liste tam olmayabilir; bilinmeyen kod geldiğinde ham
// değeri de mesaja ekliyoruz ki destek kaydında ne olduğu belli olsun.
const CODES: Record<string, string> = {
  "00": "Gönderildi",
  "01": "Gönderildi",
  "02": "Gönderildi",
  "20": "Mesaj metni hatalı veya karakter sınırı aşıldı.",
  "30":
    "Kullanıcı adı/şifre hatalı, API yetkisi yok ya da NETGSM'de IP kısıtlaması tanımlı. " +
    "IP kısıtlaması varsa kaldırın — Vercel'in sabit IP'si yoktur.",
  "40": "Gönderici başlığı NETGSM'de onaylı değil.",
  "50": "İYS kaynaklı hata — alıcının ticari ileti onayı bulunmuyor.",
  "51": "İYS marka bilgisi hatalı veya tanımsız.",
  "60": "Kayıt bulunamadı.",
  "70": "Hatalı parametre gönderildi.",
  "80": "Gönderim sınırı aşıldı.",
  "85": "Mükerrer gönderim sınırı aşıldı — aynı numaraya çok fazla tekrar.",
};

const isSuccess = (code: string) => code === "00" || code === "01" || code === "02";
const codeMessage = (code: string) =>
  CODES[code] || `NETGSM bilinmeyen yanıt kodu: ${code}`;

/**
 * İYS (İleti Yönetim Sistemi) filtresi — NETGSM'in mesajı nasıl değerlendireceğini
 * belirler. Gönderilmezse NETGSM mesajı ticari sayıp marka kaydı arar ve İYS
 * tanımlı değilse 51 koduyla reddeder.
 *
 *  "0"  — Bilgilendirme. Kargo, sipariş durumu, randevu gibi mevcut alışveriş
 *         ilişkisine dair mesajlar. İYS kontrolü yapılmaz.
 *  "11" — Ticari ileti, alıcı BİREYSEL. İYS'de onayı olmayan numaraya gitmez.
 *  "12" — Ticari ileti, alıcı TACİR. İYS'de onayı olmayan numaraya gitmez.
 *
 * Kampanya/tanıtım mesajını "0" ile göndermek İYS mevzuatına aykırıdır.
 */
export type IysFilter = "0" | "11" | "12";

export interface SendResult {
  ok: boolean;
  /** NETGSM iş kimliği — teslim raporu sorgulamak için. */
  jobId?: string;
  code?: string;
  error?: string;
  /** Numarası geçersiz olduğu için hiç denenmeyenler. */
  invalid: string[];
  /** Gönderime giren, tekilleştirilmiş normalize numaralar. */
  sent: string[];
}

/**
 * Aynı mesajı birden fazla numaraya gönderir.
 * Geçersiz numaralar sessizce yutulmaz — sonuçta `invalid` içinde döner.
 */
export async function sendSms(
  numbers: string[],
  message: string,
  iysfilter: IysFilter = "0"
): Promise<SendResult> {
  const invalid: string[] = [];
  const sent: string[] = [];
  const seen = new Set<string>();

  for (const n of numbers) {
    const p = normalizePhone(n);
    if (!p) {
      invalid.push(n);
      continue;
    }
    if (seen.has(p)) continue; // aynı numaraya iki kez göndermeyelim
    seen.add(p);
    sent.push(p);
  }

  if (!smsConfigured()) {
    return {
      ok: false,
      error:
        "NETGSM bilgileri tanımlı değil. Vercel → Settings → Environment Variables " +
        "içine NETGSM_USERCODE, NETGSM_PASSWORD ve NETGSM_HEADER girin.",
      invalid,
      sent: [],
    };
  }

  const text = String(message || "").trim();
  if (!text) {
    return { ok: false, error: "Mesaj boş olamaz.", invalid, sent: [] };
  }
  if (!sent.length) {
    return {
      ok: false,
      error: "Gönderilebilecek geçerli numara yok.",
      invalid,
      sent: [],
    };
  }

  const { encoding } = smsSegments(text);
  const auth = Buffer.from(
    `${process.env.NETGSM_USERCODE}:${process.env.NETGSM_PASSWORD}`
  ).toString("base64");

  try {
    const res = await fetch(ENDPOINT, {
      method: "POST",
      headers: {
        Authorization: `Basic ${auth}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        msgheader: process.env.NETGSM_HEADER,
        encoding,
        iysfilter,
        messages: sent.map((no) => ({ msg: text, no })),
      }),
      // NETGSM zaman zaman yavaş yanıt veriyor; isteği asılı bırakmayalım.
      signal: AbortSignal.timeout(20000),
    });

    const raw = (await res.text()).trim();

    // Yanıt JSON ({"code":"00","jobid":"..."}) veya düz metin ("00 123456")
    // gelebiliyor; ikisini de karşılıyoruz.
    let code = "";
    let jobId: string | undefined;
    try {
      const j = JSON.parse(raw) as {
        code?: string | number;
        jobid?: string | number;
      };
      code = String(j.code ?? "").trim();
      if (j.jobid != null) jobId = String(j.jobid);
    } catch {
      const parts = raw.split(/\s+/);
      code = parts[0] || "";
      jobId = parts[1];
    }

    if (!code) {
      return {
        ok: false,
        error: `NETGSM'den beklenmeyen yanıt (HTTP ${res.status}): ${raw.slice(0, 200)}`,
        invalid,
        sent,
      };
    }
    if (!isSuccess(code)) {
      return { ok: false, code, error: codeMessage(code), invalid, sent };
    }
    return { ok: true, code, jobId, invalid, sent };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return { ok: false, error: `NETGSM'e bağlanılamadı: ${msg}`, invalid, sent };
  }
}
