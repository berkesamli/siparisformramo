// Sipariş metinleri (WhatsApp) ve opsiyonel e-posta (müşteriye PDF).
import nodemailer from "nodemailer";
import type { SavedOrder, OrderItem } from "./orders";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export interface DealerBrand {
  name: string;
  phone: string;
  website?: string;
}

export function itemText(it: OrderItem, i: number): string {
  const parts = [
    `${it.artWidth}${it.artWidthUnit} x ${it.artHeight}${it.artHeightUnit}`,
    `Çerçeve: ${it.frameCode}`,
  ];
  if (it.matType !== "Paspartu Yok") {
    let m = `Paspartu: ${it.matType}`;
    if (it.matColor && it.matColor !== "-") m += ` ${it.matColor}`;
    if (it.doubleMat) m += ` + İç: ${it.innerMatType} ${it.innerMatColor}`;
    if (it.zeminEnabled) m += ` | Zemin: ${it.zeminType} ${it.zeminColor}`;
    parts.push(m);
  }
  if (it.glassType !== "Cam Yok") parts.push(`Cam: ${it.glassType}`);
  if (it.printType && it.printType !== "Baskı Yok") parts.push(`Baskı: ${it.printType}`);
  return `${i + 1}) ${parts.join(" | ")} — ${fmt(it.itemTotal)} TL`;
}

/** Müşteriye WhatsApp'tan gönderilecek sipariş özeti + takip linki. */
export function orderWhatsAppText(o: SavedOrder, brand: DealerBrand, trackUrl: string): string {
  return [
    `*${brand.name} — Sipariş ${o.orderId}*`,
    `Sayın ${o.customerName}, siparişiniz alındı.`,
    o.deliveryDate ? `Tahmini teslim: ${o.deliveryDate}` : "",
    "",
    ...o.items.map(itemText),
    "",
    o.discount > 0 ? `İndirim: -${fmt(o.discount)} TL` : "",
    `*GENEL TOPLAM: ${fmt(o.total)} TL*`,
    "",
    `Sipariş takibi: ${trackUrl}`,
    `${brand.name} | ${brand.phone}${brand.website ? " | " + brand.website : ""}`,
  ]
    .filter((l) => l !== "")
    .join("\n");
}

export function normalizePhoneWa(phone: string): string {
  const digits = String(phone || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("90") && digits.length === 12) return digits;
  if (digits.startsWith("0")) return "9" + digits;
  if (digits.startsWith("5") && digits.length === 10) return "90" + digits;
  return digits;
}

function esc(s: string): string {
  return String(s).replace(/[&<>"']/g, (m) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m] as string)
  );
}

/** SMTP tanımlıysa ve müşterinin e-postası varsa sipariş özetini (PDF ekli) yollar. */
export async function sendOrderEmail(
  o: SavedOrder,
  brand: DealerBrand,
  trackUrl: string,
  pdf?: Buffer
): Promise<boolean> {
  const host = process.env.SMTP_HOST;
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  if (!host || !user || !pass || !o.customerEmail) return false;

  const transporter = nodemailer.createTransport({
    host,
    port: Number(process.env.SMTP_PORT || 465),
    secure: (process.env.SMTP_SECURE ?? "true") !== "false",
    auth: { user, pass },
  });

  const rows = o.items
    .map((it, i) => `<li>${esc(itemText(it, i))}</li>`)
    .join("");

  await transporter.sendMail({
    from: process.env.SMTP_FROM || user,
    to: o.customerEmail,
    subject: `${brand.name} — Sipariş ${o.orderId}`,
    text: orderWhatsAppText(o, brand, trackUrl).replace(/\*/g, ""),
    html: `<div style="font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#222">
      <h2 style="margin:0 0 8px">${esc(brand.name)} — Sipariş ${esc(o.orderId)}</h2>
      <p>Sayın ${esc(o.customerName)}, siparişiniz alındı.${o.deliveryDate ? ` Tahmini teslim: <b>${esc(o.deliveryDate)}</b>` : ""}</p>
      <ul>${rows}</ul>
      ${o.discount > 0 ? `<p>İndirim: -${fmt(o.discount)} TL</p>` : ""}
      <p style="font-size:16px"><b>Genel Toplam: ${fmt(o.total)} TL</b></p>
      <p><a href="${esc(trackUrl)}">Sipariş takibi</a></p>
      <p style="color:#666">${esc(brand.name)} · ${esc(brand.phone)}</p>
    </div>`,
    attachments: pdf
      ? [{ filename: `siparis_${o.orderId}.pdf`, content: pdf, contentType: "application/pdf" }]
      : undefined,
  });
  return true;
}
