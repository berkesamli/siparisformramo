import nodemailer from "nodemailer";

export interface OrderLine {
  name: string;
  unitText: string;
  unitPriceTL: number;
  lineTotal: number;
}

export interface OrderPayload {
  orderId: string;
  employee: string;
  customer: string;
  note: string;
  rate: number;
  euroRate: number;
  discountPct: number;
  vatApplied: boolean;
  lines: OrderLine[];
  gross: number;
  discount: number;
  vatAmount: number;
  net: number;
  dateStr: string;
}

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export function orderText(o: OrderPayload): string {
  const lines = o.lines
    .map(
      (l, i) =>
        `${i + 1}. ${l.name} — ${l.unitText} — ₺${fmt(l.unitPriceTL)} → ₺${fmt(l.lineTotal)}`
    )
    .join("\n");
  return [
    `🧾 YENİ SİPARİŞ ${o.orderId}`,
    `Tarih: ${o.dateStr}`,
    `Çalışan: ${o.employee}`,
    `Müşteri: ${o.customer}`,
    o.note ? `Not: ${o.note}` : "",
    o.rate ? `Kur: ${fmt(o.rate)} TL/USD${o.euroRate ? ` | ${fmt(o.euroRate)} TL/EUR` : ""}` : "",
    "",
    lines,
    "",
    `Ara Toplam: ₺${fmt(o.gross)}`,
    `İskonto (%${o.discountPct}): ₺${fmt(o.discount)}`,
    `KDV: ${o.vatApplied ? `%20 — ₺${fmt(o.vatAmount)}` : "Uygulanmadı"}`,
    `GENEL TOPLAM: ₺${fmt(o.net)}`,
  ]
    .filter((l) => l !== "")
    .join("\n");
}

function esc(s: string): string {
  return s.replace(/[&<>"']/g, (m) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m] as string)
  );
}

export function orderHtml(o: OrderPayload): string {
  return `
  <div style="font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#222">
    <h2 style="color:#8b6914;margin:0 0 8px">Olga Çerçeve — Yeni Sipariş</h2>
    <b>Sipariş No:</b> ${o.orderId}<br>
    <b>Tarih:</b> ${o.dateStr}<br>
    <b>Çalışan:</b> ${esc(o.employee)}<br>
    <b>Müşteri:</b> ${esc(o.customer)}<br>
    <b>Not:</b> ${esc(o.note || "-")}<br>
    ${o.rate ? `<b>Kurlar:</b> ${fmt(o.rate)} TL/USD | ${fmt(o.euroRate || 0)} TL/EUR<br>` : ""}
    <br>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse">
      <tr style="background:#f7f3ef"><th>#</th><th>Ürün</th><th>Birim</th><th>Birim Fiyat (₺)</th><th>Tutar (₺)</th></tr>
      ${o.lines
        .map(
          (l, i) => `<tr>
        <td>${i + 1}</td><td>${esc(l.name)}</td><td>${esc(l.unitText)}</td>
        <td>₺ ${fmt(l.unitPriceTL)}</td><td>₺ ${fmt(l.lineTotal)}</td>
      </tr>`
        )
        .join("")}
    </table><br>
    <b>Ara Toplam:</b> ₺ ${fmt(o.gross)}<br>
    <b>İskonto:</b> %${o.discountPct} — ₺ ${fmt(o.discount)}<br>
    <b>KDV:</b> ${o.vatApplied ? "%20 — ₺ " + fmt(o.vatAmount) : "Uygulanmadı"}<br>
    <b style="font-size:16px">Genel Toplam: ₺ ${fmt(o.net)}</b>
  </div>`;
}

/** SMTP env değişkenleri tanımlıysa e-posta gönderir. */
export async function sendOrderEmail(o: OrderPayload): Promise<boolean> {
  const host = process.env.SMTP_HOST;
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  const to = process.env.ORDER_EMAIL_TO || "olgacercevee@gmail.com";
  if (!host || !user || !pass) return false;

  const transporter = nodemailer.createTransport({
    host,
    port: Number(process.env.SMTP_PORT || 465),
    secure: (process.env.SMTP_SECURE ?? "true") !== "false",
    auth: { user, pass },
  });

  await transporter.sendMail({
    from: process.env.SMTP_FROM || user,
    to,
    subject: `Yeni Sipariş ${o.orderId} — ${o.customer} — ₺ ${fmt(o.net)}`,
    text: orderText(o),
    html: orderHtml(o),
  });
  return true;
}

/**
 * WhatsApp Cloud API yapılandırıldıysa (WHATSAPP_TOKEN + WHATSAPP_PHONE_ID +
 * WHATSAPP_TO) sipariş metnini doğrudan gönderir.
 */
export async function sendOrderWhatsApp(o: OrderPayload): Promise<boolean> {
  const token = process.env.WHATSAPP_TOKEN;
  const phoneId = process.env.WHATSAPP_PHONE_ID;
  const to = process.env.WHATSAPP_TO;
  if (!token || !phoneId || !to) return false;

  const res = await fetch(
    `https://graph.facebook.com/v20.0/${phoneId}/messages`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        messaging_product: "whatsapp",
        to,
        type: "text",
        text: { body: orderText(o) },
      }),
    }
  );
  return res.ok;
}

/** Cloud API yoksa kullanılabilecek wa.me linki üretir. */
export function waLink(o: OrderPayload): string {
  const to = (process.env.WHATSAPP_TO || "908503057545").replace(/\D/g, "");
  return `https://wa.me/${to}?text=${encodeURIComponent(orderText(o))}`;
}
