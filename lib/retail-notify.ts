// Perakende sipariş bildirimleri (e-posta). SMTP ayarları lib/notify.ts ile aynıdır.

import nodemailer from "nodemailer";
import type { SavedRetailOrder, RetailItem } from "./retail-orders";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

function itemText(it: RetailItem, i: number): string {
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
  return `${i + 1}) ${parts.join(" | ")} — ₺${fmt(it.itemTotal)}`;
}

export function retailOrderText(o: SavedRetailOrder): string {
  return [
    `🖼️ PERAKENDE SİPARİŞ ${o.orderId}`,
    `Müşteri: ${o.customerName} | ${o.customerPhone}`,
    o.customerEmail ? `E-posta: ${o.customerEmail}` : "",
    `Teslim: ${o.deliveryDate || "-"}`,
    o.notes ? `Not: ${o.notes}` : "",
    "",
    ...o.items.map(itemText),
    "",
    o.discount > 0 ? `İndirim: -₺${fmt(o.discount)}` : "",
    `GENEL TOPLAM: ₺${fmt(o.total)}`,
  ]
    .filter((l) => l !== "")
    .join("\n");
}

function esc(s: string): string {
  return String(s).replace(/[&<>"']/g, (m) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m] as string)
  );
}

export function retailOrderHtml(o: SavedRetailOrder): string {
  const rows = o.items
    .map((it, i) => {
      let detay = "";
      if (it.matType !== "Paspartu Yok") {
        detay += `Paspartu: ${esc(it.matType)} ${esc(it.matColor)}`;
        if (it.doubleMat)
          detay += ` + İç: ${esc(it.innerMatType)} ${esc(it.innerMatColor)} (montaj ${esc(it.altMontaj)}mm)`;
        if (it.zeminEnabled)
          detay += ` | Zemin: ${esc(it.zeminType)} ${esc(it.zeminColor)}`;
        detay += `<br>Kenarlar (mm): üst ${it.matTop} / sağ ${it.matRight} / alt ${it.matBottom} / sol ${it.matLeft}<br>`;
      }
      if (it.glassType !== "Cam Yok") detay += `Cam: ${esc(it.glassType)}<br>`;
      if (it.printType !== "Baskı Yok") detay += `Baskı: ${esc(it.printType)}<br>`;
      return `<tr>
        <td>${i + 1}</td>
        <td>${it.artWidth} ${it.artWidthUnit} × ${it.artHeight} ${it.artHeightUnit}</td>
        <td>${esc(it.frameCode)}</td>
        <td style="font-size:12px">${detay || "-"}</td>
        <td align="right">₺ ${fmt(it.itemTotal)}</td>
      </tr>`;
    })
    .join("");

  return `
  <div style="font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#222">
    <h2 style="color:#8b6914;margin:0 0 8px">Olga Çerçeve — Perakende Sipariş</h2>
    <b>Sipariş No:</b> ${esc(o.orderId)}<br>
    <b>Personel:</b> ${esc(o.employee)}<br>
    <b>Müşteri:</b> ${esc(o.customerName)} — ${esc(o.customerPhone)}<br>
    ${o.customerEmail ? `<b>E-posta:</b> ${esc(o.customerEmail)}<br>` : ""}
    <b>Teslim Tarihi:</b> ${esc(o.deliveryDate || "-")}<br>
    <b>Not:</b> ${esc(o.notes || "-")}<br><br>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse">
      <tr style="background:#f7f3ef"><th>#</th><th>Ölçü</th><th>Çerçeve</th><th>Detay</th><th>Tutar</th></tr>
      ${rows}
    </table><br>
    <b>Ara Toplam:</b> ₺ ${fmt(o.gross)}<br>
    ${o.discount > 0 ? `<b>İndirim:</b> -₺ ${fmt(o.discount)}<br>` : ""}
    <b style="font-size:16px">Genel Toplam: ₺ ${fmt(o.total)}</b>
  </div>`;
}

/** SMTP tanımlıysa siparişi mağazaya (ve varsa müşteriye kopya) e-postalar. */
export async function sendRetailOrderEmail(o: SavedRetailOrder): Promise<boolean> {
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
    cc: o.customerEmail || undefined,
    subject: `Perakende Sipariş ${o.orderId} — ${o.customerName} — ₺ ${fmt(o.total)}`,
    text: retailOrderText(o),
    html: retailOrderHtml(o),
  });
  return true;
}
