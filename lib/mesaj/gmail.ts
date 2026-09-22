// Gmail bağlayıcısı: IMAP ile gelen kutusu okunur (imapflow), yanıt SMTP ile
// aynı hesaptan gider (nodemailer). Her hesap için Google "uygulama şifresi"
// gerekir (2 adımlı doğrulama açık olmalı).
//
//   GMAIL_HESAPLAR="olgacercevee@gmail.com:abcd efgh ijkl mnop;stasresimasmasistemleri@gmail.com:…"
//
// Konuşma anahtarı gönderenin e-posta adresidir (hesap başına). Senkron,
// gelen kutusu açılınca ve arka planda (senk=1) tetiklenir; son okunan UID
// mesaj_senk tablosunda tutulur (UIDVALIDITY değişirse baştan başlanır).

import { konusmaBul, konusmaBulVeyaOlustur, konusmaGuncelle, mesajEkle, senkOku, senkYaz } from "./db";
import { ekKaydet, ekUrl, ekTuru, EK_AZAMI_BAYT } from "./ek";
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

/** Alıntılanmış eski yazışmayı (">" satırları, "… yazdı:" bloğu) kırpar. */
export function alintiKirp(metin: string): string {
  const satirlar = String(metin || "").split("\n");
  const out: string[] = [];
  for (const s of satirlar) {
    if (/^\s*>/.test(s)) break;
    if (/^(On .+ wrote:|.+ tarihinde .+ yazdı:|-----\s*Original Message\s*-----|_{10,})\s*$/i.test(s.trim())) break;
    out.push(s);
  }
  const k = out.join("\n").trim();
  return k || String(metin || "").trim();
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
      const m = (e as Error)?.message || String(e);
      out.push({ adres: h.adres, ok: false, hata: /AUTHENTICATIONFAILED|Invalid credentials|535|Application-specific password/i.test(m) ? "Giriş reddedildi: uygulama şifresi yanlış ya da 2 adımlı doğrulama kapalı. (" + m.slice(0, 120) + ")" : m.slice(0, 200) });
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
      const hata = (e as Error)?.message || String(e);
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
      if (uidler.length > AZAMI_MESAJ) uidler = uidler.slice(-AZAMI_MESAJ);
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
    try {
      if (await mesajIsle(h, m, uidValidity, bizim)) yeni++;
    } catch (e) {
      console.warn(`Gmail mesajı işlenemedi (${h.adres} uid ${m.uid}):`, (e as Error)?.message);
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
    konusmaId: k.id, yon: "gelen", govde: konuDegisti ? `Konu: ${konu}\n\n${metin}` : metin, ekler,
    disId: messageId || `uid:${uidValidity}:${m.uid}`, gonderen: ad, at, sessiz: otomatik,
  });
  return r.yeni;
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
