// 150×100 mm YATAY kargo etiketi (pdfkit).
// "etiket" uygulamasındaki html2canvas+jsPDF üretiminin sunucu tarafı karşılığı:
// tek renk baskıya uygun, siyah-beyaz, gönderici şubeye göre değişen başlık.

import path from "path";
import PDFDocument from "pdfkit";
import { branchInfo, customerTitle, type Customer } from "./customers";

const FONT = path.join(process.cwd(), "assets", "DejaVuSans.ttf");
const FONT_BOLD = path.join(process.cwd(), "assets", "DejaVuSans-Bold.ttf");
const LOGO = path.join(process.cwd(), "assets", "olga-logo.png");

const MM = 2.83465;
const mm = (v: number) => v * MM;

const PAGE_W = mm(150);
const PAGE_H = mm(100);
const PAD = mm(6);

export function generateLabelPdf(c: Customer, count = 1): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: [PAGE_W, PAGE_H],
      margin: 0,
      font: FONT,
    });
    const chunks: Buffer[] = [];
    doc.on("data", (b: Buffer) => chunks.push(b));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
    doc.page.margins = { top: 0, bottom: 0, left: 0, right: 0 };

    const pages = Math.max(1, Math.min(50, count));
    for (let i = 0; i < pages; i++) {
      if (i > 0) doc.addPage({ size: [PAGE_W, PAGE_H], margin: 0 });
      drawLabel(doc, c);
    }
    doc.end();
  });
}

function drawLabel(doc: PDFKit.PDFDocument, c: Customer) {
  const b = branchInfo(c.branch);
  const innerW = PAGE_W - PAD * 2;

  // Dış çerçeve
  doc
    .rect(mm(1.5), mm(1.5), PAGE_W - mm(3), PAGE_H - mm(3))
    .lineWidth(1.2)
    .strokeColor("#000")
    .stroke();

  // ===== Gönderici başlığı: logo + şube bilgileri =====
  const headTop = PAD;
  try {
    doc.image(LOGO, PAD, headTop, { fit: [mm(22), mm(20)] });
  } catch {
    /* logo yoksa yazıyla devam */
  }
  const firmX = PAD + mm(26);
  doc.font(FONT_BOLD).fontSize(13.5).fillColor("#000");
  doc.text(b.name, firmX, headTop, { width: innerW - mm(26) });
  doc.font(FONT_BOLD).fontSize(10.5);
  doc.text(b.cityTel, firmX, headTop + mm(6), { width: innerW - mm(26) });
  // Adres 1 veya 2 satır olabilir (İstanbul şubesi uzun) — web sitesi satırı
  // adresin gerçek yüksekliğine göre konumlanır ki üst üste binmesin.
  doc.font(FONT).fontSize(9.5).fillColor("#222");
  const addrW = innerW - mm(26);
  const addrTop = headTop + mm(10.5);
  const addrText = `${b.addr1}, ${b.addr2}`;
  const addrH = doc.heightOfString(addrText, { width: addrW });
  doc.text(addrText, firmX, addrTop, { width: addrW });
  const webTop = addrTop + addrH + mm(0.6);
  doc.text(b.website, firmX, webTop, { width: addrW });
  const webH = doc.heightOfString(b.website, { width: addrW });

  // Başlık altı kalın çizgi — metnin ve logonun altında kalacak şekilde
  const sepY = Math.max(webTop + webH, headTop + mm(19)) + mm(2.5);
  doc
    .moveTo(PAD, sepY)
    .lineTo(PAGE_W - PAD, sepY)
    .lineWidth(1.4)
    .strokeColor("#000")
    .stroke();

  // ===== Alıcı kutusu =====
  const boxTop = sepY + mm(5);
  const boxH = PAGE_H - boxTop - PAD;
  doc.rect(PAD, boxTop, innerW, boxH).lineWidth(1).strokeColor("#000").stroke();

  const tx = PAD + mm(4);
  const tw = innerW - mm(8);
  let y = boxTop + mm(3.5);

  doc.font(FONT_BOLD).fontSize(12).fillColor("#000");
  doc.text("Alıcı", tx, y);
  y += mm(6.5);

  doc.font(FONT_BOLD).fontSize(13);
  doc.text(customerTitle(c), tx, y, { width: tw, height: mm(6), ellipsis: true });
  y += mm(7);

  const contact = [c.phone && `Tel: ${c.phone}`, c.email && `E-posta: ${c.email}`]
    .filter(Boolean)
    .join("  •  ");
  if (contact) {
    doc.font(FONT).fontSize(10).fillColor("#222");
    doc.text(contact, tx, y, { width: tw, height: mm(5), ellipsis: true });
    y += mm(6);
  }

  // Adres satırları
  const lines: string[] = [];
  if (c.addr1) lines.push(c.addr1);
  if (c.addr2) lines.push(c.addr2);
  const cityLine = [c.district, c.city].filter(Boolean).join(" / ");
  if (cityLine) lines.push(cityLine);
  const pcCountry = [c.postalCode, c.country].filter(Boolean).join(" ");
  if (pcCountry) lines.push(pcCountry);

  doc.font(FONT).fontSize(11.5).fillColor("#000");
  doc.text(lines.length ? lines.join("\n") : "-", tx, y + mm(1), {
    width: tw,
    lineGap: 1.5,
  });

  // Not (varsa) — kutunun altına küçük punto
  if (c.note) {
    doc.font(FONT).fontSize(8.5).fillColor("#333");
    doc.text(c.note, tx, boxTop + boxH - mm(6), {
      width: tw,
      height: mm(5),
      ellipsis: true,
    });
  }
}
