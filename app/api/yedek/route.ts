// Gecelik otomatik yedek — vercel.json'daki cron her gece 02:00 UTC'de
// (İstanbul 05:00) çağırır. Blob'daki TÜM veri dosyaları tek bir gzip
// paketine toplanır: backups/YYYY-MM-DD.json.gz. Paket ayrıca e-postayla
// gönderilir (SMTP tanımlıysa) — böylece Blob deposu tamamen kaybolsa bile
// posta kutusunda depo dışı bir kopya durur. 14 günden eski yedekler silinir.
//
// Güvenlik: Vercel cron istekleri Authorization: Bearer <CRON_SECRET> başlığı
// taşır (CRON_SECRET ortam değişkeni tanımlandığında Vercel bunu otomatik
// ekler). Panelden elle tetikleme için firma sahibi oturumu da kabul edilir.

import { NextResponse } from "next/server";
import { gzipSync } from "zlib";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import { blobConfigured, istanbulDateKey } from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

const SAKLANACAK_GUN = 14;
const EK_LIMIT_BYTE = 10 * 1024 * 1024; // e-posta ekine sığacak azami boyut

async function yetkili(req: Request): Promise<boolean> {
  const sir = process.env.CRON_SECRET;
  const auth = req.headers.get("authorization") || "";
  if (sir && auth === `Bearer ${sir}`) return true;
  const user = await getSessionUser();
  return !!user && user.role === "staff" && isOwner(user.username);
}

export async function GET(req: Request) {
  if (!(await yetkili(req))) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  if (!blobConfigured()) {
    return NextResponse.json(
      { ok: false, error: "Blob deposu yapılandırılmamış." },
      { status: 503 }
    );
  }

  const { list, get, put, del } = await import("@vercel/blob");
  const bugun = istanbulDateKey();

  // 1) Tüm veri dosyalarını listele (yedekler hariç, sayfa sayfa)
  const yollar: string[] = [];
  let cursor: string | undefined;
  do {
    const r = await list({ limit: 1000, cursor });
    for (const b of r.blobs) {
      if (b.pathname.startsWith("backups/")) continue;
      yollar.push(b.pathname);
    }
    cursor = r.hasMore ? r.cursor : undefined;
  } while (cursor);

  // 2) İçerikleri oku — 25'erli gruplar halinde (hız/istek dengesi).
  // Dosyalar ham metin olarak saklanır: geri yükleme birebir yazmaktan ibaret.
  const dosyalar: Record<string, string> = {};
  const okunamadi: string[] = [];
  for (let i = 0; i < yollar.length; i += 25) {
    await Promise.all(
      yollar.slice(i, i + 25).map(async (yol) => {
        try {
          const r = await get(yol, { access: "private", useCache: false });
          if (!r || r.statusCode !== 200 || !r.stream) {
            okunamadi.push(yol);
            return;
          }
          dosyalar[yol] = await new Response(r.stream).text();
        } catch {
          okunamadi.push(yol);
        }
      })
    );
  }

  // 3) Paketle, sıkıştır, backups/ altına yaz
  const paket = JSON.stringify({
    olusturuldu: new Date().toISOString(),
    gun: bugun,
    dosyaSayisi: Object.keys(dosyalar).length,
    okunamadi,
    dosyalar,
  });
  const gz = gzipSync(Buffer.from(paket, "utf8"));
  const yedekYolu = `backups/${bugun}.json.gz`;
  await put(yedekYolu, gz, {
    access: "private",
    contentType: "application/gzip",
    addRandomSuffix: false,
    allowOverwrite: true,
  });

  // 4) Eski yedekleri temizle (14 günden eski)
  const sinir = new Date(Date.now() - SAKLANACAK_GUN * 24 * 60 * 60 * 1000)
    .toISOString()
    .slice(0, 10);
  const silinen: string[] = [];
  try {
    const eski = await list({ prefix: "backups/", limit: 1000 });
    for (const b of eski.blobs) {
      const m = b.pathname.match(/^backups\/(\d{4}-\d{2}-\d{2})\.json\.gz$/);
      if (m && m[1] < sinir) {
        try {
          await del(b.pathname);
          silinen.push(b.pathname);
        } catch {
          /* silinemeyen eski yedek bir sonraki gece tekrar denenir */
        }
      }
    }
  } catch {
    /* temizlik başarısız olsa da yedek alınmış durumda */
  }

  // 5) Depo dışı kopya: yedeği e-postayla gönder (SMTP tanımlıysa)
  let emailSent = false;
  try {
    const host = process.env.SMTP_HOST;
    const smtpUser = process.env.SMTP_USER;
    const pass = process.env.SMTP_PASS;
    const to =
      process.env.BACKUP_EMAIL_TO ||
      process.env.ORDER_EMAIL_TO ||
      "olgacercevee@gmail.com";
    if (host && smtpUser && pass) {
      const nodemailer = (await import("nodemailer")).default;
      const transporter = nodemailer.createTransport({
        host,
        port: Number(process.env.SMTP_PORT || 465),
        secure: (process.env.SMTP_SECURE ?? "true") !== "false",
        auth: { user: smtpUser, pass },
      });
      const mb = (n: number) => (n / (1024 * 1024)).toFixed(2);
      await transporter.sendMail({
        from: process.env.SMTP_FROM || smtpUser,
        to,
        subject: `Olga Sipariş — günlük yedek ${bugun} (${Object.keys(dosyalar).length} dosya)`,
        text:
          `Günlük yedek alındı.\n\n` +
          `Tarih: ${bugun}\nDosya sayısı: ${Object.keys(dosyalar).length}\n` +
          `Okunamayan: ${okunamadi.length}\n` +
          `Paket boyutu: ${mb(gz.length)} MB (sıkıştırılmış)\n` +
          `Depodaki yol: ${yedekYolu}\n` +
          (gz.length <= EK_LIMIT_BYTE
            ? `\nYedek dosyası ektedir — bu e-posta depo dışı kopyanızdır.`
            : `\nYedek ${mb(EK_LIMIT_BYTE)} MB ek sınırını aştığı için ekte değil; Blob deposundaki kopya kullanılmalıdır.`),
        attachments:
          gz.length <= EK_LIMIT_BYTE
            ? [
                {
                  filename: `olga-yedek-${bugun}.json.gz`,
                  content: gz,
                  contentType: "application/gzip",
                },
              ]
            : undefined,
      });
      emailSent = true;
    }
  } catch (err) {
    console.error("Yedek e-postası gönderilemedi:", err);
  }

  return NextResponse.json({
    ok: true,
    gun: bugun,
    dosyaSayisi: Object.keys(dosyalar).length,
    okunamadi: okunamadi.length,
    boyutByte: gz.length,
    yol: yedekYolu,
    silinenEskiYedek: silinen.length,
    emailSent,
  });
}
