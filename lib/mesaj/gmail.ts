// Gmail bağlayıcısı: IMAP ile gelen kutusu okunur (imapflow), yanıt SMTP ile
// aynı hesaptan gider (nodemailer). Her hesap için Google "uygulama şifresi"
// gerekir (2 adımlı doğrulama açık olmalı).
//
//   GMAIL_HESAPLAR="olgacercevee@gmail.com:abcd efgh ijkl mnop;stasresimasmasistemleri@gmail.com:…"
//
// Konuşma anahtarı gönderenin e-posta adresidir (hesap başına). Senkron,
// gelen kutusu açılınca ve arka planda (senk=1) tetiklenir; son okunan UID
// mesaj_senk tablosunda tutulur (UIDVALIDITY değişirse baştan başlanır).

import { htmlEksikMesajlar, konusmaBul, konusmaBulVeyaOlustur, konusmaGuncelle, mesajEkle, mesajHtmlYaz, senkOku, senkYaz } from "./db";
import { ekKaydet, ekUrl, ekTuru, EK_AZAMI_BAYT } from "./ek";
import { HTML_AZAMI, htmlTemizle } from "./eposta-html";
import { musteriEsle } from "./musteri-esle";
import type { Ek, Konusma } from "./tur";

export interface GmailHesap { adres: string; sifre: string }

const ILK_SENK_GUN = 7;
const SENK_ARALIK_MS = 45_000;
const AZAMI_MESAJ = 60;

export function gmailHesaplar(): GmailHesap[] {
  const raw = process.env.GMAIL_HESAPLAR || "";
  const out: GmailHesap[] = [];
  for (const parca of raw.split(/[;\n]+/)) {
    const p = parca.trim();
    if (!p) continue;
    const i = p.indexOf(":");
    if (i < 1) continue;
    const adres = p.slice(0, i).trim().toLowerCase().replace(/^["']|["']$/g, "");
    // Google uygulama şifresi 4'lü gruplarla gösterilir ("abcd efgh …"); boşluklar atılır.
    const sifre = p.slice(i + 1).trim().replace(/^["']|["']$/g, "").replace(/\s+/g, "");
    if (adres.includes("@") && sifre) out.push({ adres, sifre });
  }
  return out;
}

export function gmailConfigured(): boolean {
  return gmailHesaplar().length > 0;
}

/** Bizim taraf sayılan adresler: Gmail hesapları + sipariş bildirimi gönderen SMTP adresi. */
function bizimAdresler(): Set<string> {
  const s = new Set(gmailHesaplar().map((h) => h.adres));
  for (const v of [process.env.SMTP_USER, process.env.SMTP_FROM]) {
    const m = String(v || "").match(/[\w.+-]+@[\w.-]+/);
    if (m) s.add(m[0].toLowerCase());
  }
  return s;
}

const OTOMATIK = /no-?reply|do-?not-?reply|mailer-daemon|postmaster|notification|newsletter|bildirim|bulten|noreply|info@.*(google|facebook|meta|instagram)\.com/i;

/** HTML gövdeyi düz metne indirger (metin parçası yoksa). */
export function htmlToText(html: string): string {
  return String(html || "")
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h[1-6]|blockquote)>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .split("\n").map((l) => l.replace(/[ \t]+/g, " ").trim()).join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/**
 * Alıntılanmış eski yazışmayı kırpar: Outlook ayracından sonrası ve SONDAKİ ">" / "… yazdı:" bloğu atılır;
 * satır arası (alıntının altına yazılan) yanıtlar korunur.
 */
export function alintiKirp(metin: string): string {
  const satirlar = String(metin || "").split("\n");
  const ATIF = /^(On .+ wrote:|.+ tarihinde .+ yazdı:)\s*$/i;
  const AYRAC = /^(-----\s*Original Message\s*-----|_{10,})\s*$/i;
  let son = satirlar.length;
  const ayrac = satirlar.findIndex((s) => AYRAC.test(s.trim()));
  if (ayrac >= 0) son = ayrac;
  while (son > 0 && (/^\s*$/.test(satirlar[son - 1]) || /^\s*>/.test(satirlar[son - 1]) || ATIF.test(satirlar[son - 1].trim()))) son--;
  const k = satirlar.slice(0, son).join("\n").trim();
  return k || String(metin || "").trim();
}

/**
 * imapflow hataları genel "Command failed" mesajıyla gelir; Google'ın asıl yanıtı responseText'tedir
 * (örn. "[AUTHENTICATIONFAILED] Invalid credentials (Failure)"). Buradan anlaşılır Türkçe açıklama üretilir.
 */
export function imapHata(e: unknown): string {
  const err = (e || {}) as { message?: string; responseText?: string; responseStatus?: string; serverResponseCode?: string; code?: string; authenticationFailed?: boolean };
  const ham = [err.serverResponseCode, err.responseText].filter(Boolean).join(" ") || err.message || String(e);
  const kod = String(err.code || "");
  if (/AUTHENTICATIONFAILED|Invalid credentials|authenticationFailed/i.test(ham) || err.authenticationFailed) {
    return `Google girişi reddetti: uygulama şifresi yanlış ya da bu hesapta 2 adımlı doğrulama kapalı. Normal Gmail şifresi çalışmaz; https://myaccount.google.com/apppasswords adresinden 16 harfli uygulama şifresi alın. (${ham})`;
  }
  if (/Application-specific password required/i.test(ham)) return `Google uygulama şifresi istiyor: normal şifre girilmiş. https://myaccount.google.com/apppasswords (${ham})`;
  if (/Please log in via your web browser|Web login required/i.test(ham)) return `Google girişi engelledi (yeni konum). Hesaba tarayıcıdan girip https://accounts.google.com/DisplayUnlockCaptcha adresinde izin verin, birkaç dakika sonra tekrar deneyin. (${ham})`;
  if (/IMAP access is disabled|IMAP is disabled/i.test(ham)) return `Bu hesapta IMAP kapalı: Gmail → Ayarlar → Yönlendirme ve POP/IMAP → IMAP'i etkinleştir. (${ham})`;
  if (/Too many simultaneous connections|throttl|rate/i.test(ham)) return `Google geçici olarak sınırladı; birkaç dakika sonra tekrar deneyin. (${ham})`;
  if (/ETIMEDOUT|timeout|ECONNRESET|ECONNREFUSED|ENOTFOUND|EAI_AGAIN/i.test(kod + " " + ham)) return `imap.gmail.com sunucusuna ulaşılamadı (ağ/zaman aşımı). (${kod || ham})`;
  return ham.slice(0, 240);
}

/** Her hesap için IMAP girişi + INBOX sayısı (Ayarlar test kartı). Şifre asla dönmez. */
export async function gmailTest(): Promise<{ adres: string; ok: boolean; inbox?: number; okunmamis?: number; hata?: string }[]> {
  const out: { adres: string; ok: boolean; inbox?: number; okunmamis?: number; hata?: string }[] = [];
  for (const h of gmailHesaplar()) {
    try {
      const { ImapFlow } = await import("imapflow");
      const client = new ImapFlow({ host: "imap.gmail.com", port: 993, secure: true, auth: { user: h.adres, pass: h.sifre }, logger: false, connectionTimeout: 12_000, greetingTimeout: 12_000, socketTimeout: 20_000 } as any);
      await client.connect();
      try {
        const st = await client.status("INBOX", { messages: true, unseen: true });
        out.push({ adres: h.adres, ok: true, inbox: Number(st.messages) || 0, okunmamis: Number(st.unseen) || 0 });
      } finally {
        await client.logout().catch(() => {});
      }
    } catch (e) {
      out.push({ adres: h.adres, ok: false, hata: imapHata(e) });
    }
  }
  return out;
}

const senkSuruyor = new Set<string>();

export interface GmailSenkSonuc { hesap: string; yeni: number; atlandi?: boolean; hata?: string }

/** Tüm hesapları (ya da birini) senkronlar. zorla=false ise 45 sn içinde tekrar edilmez. */
export async function gmailSenk(o: { zorla?: boolean; hesap?: string } = {}): Promise<GmailSenkSonuc[]> {
  const sonuc: GmailSenkSonuc[] = [];
  for (const h of gmailHesaplar()) {
    if (o.hesap && h.adres !== o.hesap) continue;
    if (senkSuruyor.has(h.adres)) { sonuc.push({ hesap: h.adres, yeni: 0, atlandi: true }); continue; }
    senkSuruyor.add(h.adres);
    try {
      sonuc.push(await hesapSenk(h, Boolean(o.zorla)));
    } catch (e) {
      const hata = imapHata(e);
      console.error(`Gmail senkron hatası (${h.adres}):`, hata);
      sonuc.push({ hesap: h.adres, yeni: 0, hata });
    } finally {
      senkSuruyor.delete(h.adres);
    }
  }
  return sonuc;
}

interface HamMesaj { uid: number; source: Buffer; threadId?: string; internalDate?: Date }

async function hesapSenk(h: GmailHesap, zorla: boolean): Promise<GmailSenkSonuc> {
  const anahtar = `gmail:${h.adres}`;
  const senk = await senkOku(anahtar);
  if (!zorla && senk && Date.now() - new Date(senk.at).getTime() < SENK_ARALIK_MS) return { hesap: h.adres, yeni: 0, atlandi: true };

  const { ImapFlow } = await import("imapflow");
  const client = new ImapFlow({ host: "imap.gmail.com", port: 993, secure: true, auth: { user: h.adres, pass: h.sifre }, logger: false, connectionTimeout: 15_000, greetingTimeout: 15_000, socketTimeout: 60_000 } as any);
  const ham: HamMesaj[] = [];
  let uidValidity = "";
  let sonUid = 0;
  await client.connect();
  try {
    const lock = await client.getMailboxLock("INBOX");
    try {
      const mb = client.mailbox;
      if (!mb) throw new Error("INBOX açılamadı.");
      uidValidity = String(mb.uidValidity ?? "");
      const [eskiValidity, eskiUidStr] = (senk?.deger || "").split(":");
      const eskiUid = eskiValidity === uidValidity ? Number(eskiUidStr) || 0 : 0;
      sonUid = eskiUid;
      let uidler: number[] = [];
      if (eskiUid > 0) {
        const bulunan = await client.search({ uid: `${eskiUid + 1}:*` }, { uid: true });
        uidler = (Array.isArray(bulunan) ? bulunan : []).filter((u) => u > eskiUid);
      } else {
        const bulunan = await client.search({ since: new Date(Date.now() - ILK_SENK_GUN * 86400_000) }, { uid: true });
        uidler = Array.isArray(bulunan) ? bulunan : [];
      }
      uidler.sort((a, b) => a - b);
      // Artımlı senkron: en ESKİ AZAMI_MESAJ (imleç işlenen son UID'ye ilerler, kalanı sonraki tur alır).
      // İlk senkron: son 7 günün en yeni AZAMI_MESAJ'ı (bilinçli başlangıç sınırı).
      if (uidler.length > AZAMI_MESAJ) uidler = eskiUid > 0 ? uidler.slice(0, AZAMI_MESAJ) : uidler.slice(-AZAMI_MESAJ);
      if (uidler.length) {
        for await (const msg of client.fetch(uidler, { uid: true, source: true, threadId: true, internalDate: true }, { uid: true })) {
          if (msg.source) ham.push({ uid: msg.uid, source: msg.source, threadId: msg.threadId, internalDate: msg.internalDate instanceof Date ? msg.internalDate : undefined });
        }
      }
    } finally {
      lock.release();
    }
  } finally {
    await client.logout().catch(() => {});
  }

  const bizim = bizimAdresler();
  let yeni = 0;
  for (const m of ham) {
    // Geçici hata (DB/Blob zaman aşımı) için bir kez daha dene; yine olmazsa atla ki bozuk bir e-posta hesabı kilitlemesin.
    for (let deneme = 0; deneme < 2; deneme++) {
      try {
        if (await mesajIsle(h, m, uidValidity, bizim)) yeni++;
        break;
      } catch (e) {
        if (deneme === 0) { await new Promise((r) => setTimeout(r, 500)); continue; }
        console.warn(`Gmail mesajı işlenemedi, atlandı (${h.adres} uid ${m.uid}):`, (e as Error)?.message);
      }
    }
    if (m.uid > sonUid) sonUid = m.uid;
  }
  await senkYaz(anahtar, `${uidValidity}:${sonUid}`);
  return { hesap: h.adres, yeni };
}

async function mesajIsle(h: GmailHesap, m: HamMesaj, uidValidity: string, bizim: Set<string>): Promise<boolean> {
  const { simpleParser } = await import("mailparser");
  const p = await simpleParser(m.source);
  const from = p.from?.value?.[0];
  const adres = (from?.address || "").toLowerCase();
  if (!adres) return false;
  if (bizim.has(adres)) return false; // kendi gönderdiğimiz (örn. sipariş bildirimi) — atla
  const ad = from?.name || adres;
  const konu = (p.subject || "").trim() || "(konu yok)";
  const otomatik = OTOMATIK.test(adres) || /^(auto-generated|auto-replied)/i.test(String(p.headers.get("auto-submitted") || "")) || /bulk|list/i.test(String(p.headers.get("precedence") || ""));
  const metinHam = (p.text && p.text.trim()) || (p.html ? htmlToText(p.html) : "");
  const metin = alintiKirp(metinHam).slice(0, 20_000);
  const messageId = p.messageId || null;
  const refs = ([] as string[]).concat(Array.isArray(p.references) ? p.references : p.references ? [p.references] : []);
  if (messageId) refs.push(messageId);

  // Konu değiştiyse gövdenin başına yazılır (aynı kişiyle farklı konular tek konuşmada akar).
  const onceki = await konusmaBul("email", h.adres, adres);
  const konuDegisti = Boolean(onceki && onceki.baslik && onceki.baslik !== konu);
  const k = await konusmaBulVeyaOlustur({
    kanal: "email", hesap: h.adres, disKimlik: adres, ad, baslik: konu,
    meta: { eposta: adres, sonMessageId: messageId, references: refs.slice(-20), threadId: m.threadId || undefined },
  });
  if (!k.musteriId) {
    const e = await musteriEsle({ eposta: adres }).catch(() => null);
    if (e) await konusmaGuncelle(k.id, { musteriId: e.musteriId, musteriTur: e.musteriTur });
  }
  // HTML gövde: temizlenip saklanır; arayüz kum havuzlu çerçevede gösterir, düz metin (govde) yedek ve yapay zekâ taslağı için kalır.
  const html = (await htmlHazirla(p, k.id)) || undefined;
  const ekler: Ek[] = [];
  for (const a of (p.attachments || []).slice(0, 8)) {
    if (a.contentDisposition === "inline" && a.cid) continue;
    const dosya = a.filename || `ek.${(a.contentType || "bin").split("/")[1] || "bin"}`;
    const ek: Ek = { tur: ekTuru(a.contentType || "", dosya), ad: dosya, mime: a.contentType, boyut: a.size };
    if (a.content && a.size <= EK_AZAMI_BAYT) {
      const yol = await ekKaydet(k.id, dosya, a.contentType || "application/octet-stream", a.content);
      if (yol) ek.url = ekUrl(yol);
    }
    ekler.push(ek);
  }
  const at = p.date instanceof Date && !isNaN(p.date.getTime()) ? p.date : m.internalDate || new Date();
  const r = await mesajEkle({
    konusmaId: k.id, yon: "gelen", govde: konuDegisti ? `Konu: ${konu}\n\n${metin}` : metin, html, ekler,
    disId: messageId || `uid:${uidValidity}:${m.uid}`, gonderen: ad, at, sessiz: otomatik,
  });
  return r.yeni;
}

/** Ayrıştırılmış e-postanın HTML'ini temizler; gömülü (cid:) görselleri ek deposuna alıp adreslerini değiştirir. HTML yoksa/çok büyükse "". */
async function htmlHazirla(p: { html?: string | false; attachments?: any[] }, konusmaId: string): Promise<string> {
  if (!p.html) return "";
  let h = htmlTemizle(p.html);
  for (const a of p.attachments || []) {
    if (!(a.contentDisposition === "inline" && a.cid) || !a.content || !h.includes(`cid:${a.cid}`) || a.size > 3 * 1024 * 1024) continue;
    const yol = await ekKaydet(konusmaId, a.filename || "gorsel", a.contentType || "image/png", a.content);
    if (yol) h = h.split(`cid:${a.cid}`).join(ekUrl(yol));
  }
  return h && h.length <= HTML_AZAMI ? h : "";
}

const HTML_TAMAMLA_SURE_MS = 9_000;

/**
 * Geriye dönük tamamlama: HTML'i hiç okunmamış (bu özellikten önce gelen) e-postalar konuşma açılınca
 * Gmail'den Message-ID ile yeniden okunur ve HTML'i yazılır (en çok `adet` mesaj, ~9 sn bütçe).
 * "Tüm Postalar" klasöründe aranır (arşivlenmiş olsa da bulunur); HTML'i olmayan mesaja "" yazılır ki tekrar denenmesin.
 * Döner: tamamlanan mesaj sayısı. Hata durumunda sessizce 0 (konuşma yine açılır).
 */
export async function gmailHtmlTamamla(k: Konusma, adet = 3): Promise<number> {
  if (k.kanal !== "email") return 0;
  const h = gmailHesaplar().find((x) => x.adres === k.hesap);
  if (!h) return 0;
  const eksik = (await htmlEksikMesajlar(k.id, adet)).filter((m) => m.disId && m.disId.includes("@"));
  if (!eksik.length) return 0;
  const { ImapFlow } = await import("imapflow");
  const { simpleParser } = await import("mailparser");
  const client = new ImapFlow({ host: "imap.gmail.com", port: 993, secure: true, auth: { user: h.adres, pass: h.sifre }, logger: false, connectionTimeout: 8_000, greetingTimeout: 8_000, socketTimeout: 20_000 } as any);
  const baslangic = Date.now();
  let n = 0;
  const islem = async () => {
    await client.connect();
    try {
      // Gmail "Tüm Postalar" (özel kullanım \All; adı dile göre değişir), yoksa INBOX
      let kutu = "INBOX";
      try {
        const liste = await client.list();
        const hepsi = (Array.isArray(liste) ? liste : []).find((x: any) => String(x.specialUse || "").toLowerCase() === "\\all");
        if (hepsi?.path) kutu = hepsi.path;
      } catch { /* liste alınamazsa INBOX */ }
      const lock = await client.getMailboxLock(kutu);
      try {
        for (const m of eksik) {
          if (Date.now() - baslangic > HTML_TAMAMLA_SURE_MS) break;
          const bulunan = await client.search({ header: { "message-id": m.disId as string } }, { uid: true });
          const uid = Array.isArray(bulunan) && bulunan.length ? bulunan[bulunan.length - 1] : 0;
          if (!uid) { await mesajHtmlYaz(m.id, ""); continue; }
          const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
          if (!msg || !msg.source) { await mesajHtmlYaz(m.id, ""); continue; }
          const p = await simpleParser(msg.source);
          const html = await htmlHazirla(p, k.id);
          await mesajHtmlYaz(m.id, html);
          if (html) n++;
        }
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
    return n;
  };
  // Süre bütçesi: aşılırsa konuşma tamamlanan kadarıyla açılır; bağlantı arka planda kapanır
  return Promise.race([
    islem().catch((e) => { console.warn(`Gmail HTML tamamlama (${h.adres}):`, imapHata(e)); return n; }),
    new Promise<number>((r) => setTimeout(() => r(n), HTML_TAMAMLA_SURE_MS + 1_500)),
  ]);
}

function kacir(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c] as string));
}

/** Konuşmanın hesabından (SMTP) yanıt gönderir; Message-ID döner. */
export async function gmailGonder(k: Konusma, metin: string): Promise<{ disId: string; references: string[] }> {
  const h = gmailHesaplar().find((x) => x.adres === k.hesap) || gmailHesaplar()[0];
  if (!h) throw new Error("Gmail hesabı ayarlı değil (GMAIL_HESAPLAR).");
  const nodemailer = (await import("nodemailer")).default;
  const t = nodemailer.createTransport({ host: "smtp.gmail.com", port: 465, secure: true, auth: { user: h.adres, pass: h.sifre } });
  const konu = k.baslik && k.baslik !== "(konu yok)" ? (/^(re|ynt|yanıt)\s*:/i.test(k.baslik) ? k.baslik : `Re: ${k.baslik}`) : "Olga Çerçeve";
  const refs = Array.isArray(k.meta.references) ? (k.meta.references as string[]) : [];
  const inReplyTo = typeof k.meta.sonMessageId === "string" ? k.meta.sonMessageId : undefined;
  const info = await t.sendMail({
    from: `"Olga Çerçeve" <${h.adres}>`,
    to: k.disKimlik,
    subject: konu,
    text: metin,
    html: `<div style="font-family:Arial,sans-serif;font-size:14px;white-space:pre-wrap">${kacir(metin)}</div>`,
    inReplyTo,
    references: refs.length ? refs.join(" ") : undefined,
  });
  const id = info.messageId || `smtp-${Date.now()}`;
  return { disId: id, references: [...refs, id].slice(-20) };
}
