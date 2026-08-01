// Sipariş fişi PDF üretimi (pdfkit + DejaVu Sans — Türkçe karakter desteği).
// Hem e-posta ekinde hem de /api/orders/pdf indirme ucunda kullanılır.

import path from "path";
import PDFDocument from "pdfkit";

export interface PdfOrder {
  orderId: string;
  dateStr: string;
  status?: string;
  employee: string;
  customer: string;
  note: string;
  discountPct: number;
  vatApplied: boolean;
  lines: { name: string; unitText: string; unitPriceTL: number; lineTotal: number }[];
  gross: number;
  discount: number;
  vatAmount: number;
  net: number;
}

const FONT = path.join(process.cwd(), "assets", "DejaVuSans.ttf");
const FONT_BOLD = path.join(process.cwd(), "assets", "DejaVuSans-Bold.ttf");
const BRAND = "#8b6914";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export function generateOrderPdf(order: PdfOrder): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    // font seçeneği ile başlatınca pdfkit standart (Helvetica) fontlarını hiç
    // yüklemez — Türkçe karakterler ve Vercel paketi için gereklidir.
    const doc = new PDFDocument({ size: "A4", margin: 50, font: FONT });
    const chunks: Buffer[] = [];
    doc.on("data", (c: Buffer) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    const pageWidth = doc.page.width - 100; // marjlar düşülmüş

    // Başlık
    doc.font(FONT_BOLD).fontSize(20).fillColor(BRAND).text("OLGA ÇERÇEVE");
    doc.font(FONT).fontSize(10).fillColor("#666").text("Sipariş Fişi");
    doc.moveDown(0.6);

    // Üst bilgiler
    doc.fontSize(10).fillColor("#000");
    const info: [string, string][] = [
      ["Sipariş No", order.orderId],
      ["Tarih", order.dateStr],
      ["Çalışan", order.employee],
      ["Müşteri", order.customer || "-"],
    ];
    if (order.status) info.push(["Durum", order.status]);
    if (order.note) info.push(["Not", order.note]);
    for (const [k, v] of info) {
      doc.font(FONT_BOLD).text(`${k}: `, { continued: true });
      doc.font(FONT).text(v);
    }
    doc.moveDown(0.8);

    // Tablo başlığı
    const cols = [26, pageWidth * 0.34, pageWidth * 0.3, pageWidth * 0.14, pageWidth * 0.16];
    const colX: number[] = [];
    let x = 50;
    for (const w of cols) {
      colX.push(x);
      x += w;
    }
    const rowHeight = (texts: string[], widths: number[]) =>
      Math.max(
        ...texts.map((t, i) => doc.heightOfString(t, { width: widths[i] - 6 }))
      ) + 8;

    const drawRow = (
      texts: string[],
      opts: { bold?: boolean; bg?: string } = {}
    ) => {
      const h = rowHeight(texts, cols);
      if (doc.y + h > doc.page.height - 60) doc.addPage();
      const y = doc.y;
      if (opts.bg) {
        doc.rect(50, y, pageWidth, h).fill(opts.bg);
      }
      doc.fillColor("#000").font(opts.bold ? FONT_BOLD : FONT).fontSize(9);
      texts.forEach((t, i) => {
        doc.text(t, colX[i] + 3, y + 4, {
          width: cols[i] - 6,
          align: i >= 3 ? "right" : "left",
        });
      });
      doc.y = y + h;
      doc.x = 50;
      doc
        .moveTo(50, doc.y)
        .lineTo(50 + pageWidth, doc.y)
        .strokeColor("#dddddd")
        .lineWidth(0.5)
        .stroke();
    };

    drawRow(["#", "Ürün", "Birim", "B.Fiyat (₺)", "Tutar (₺)"], {
      bold: true,
      bg: "#f5f0e6",
    });
    order.lines.forEach((l, i) => {
      drawRow([String(i + 1), l.name, l.unitText, fmt(l.unitPriceTL), fmt(l.lineTotal)]);
    });

    doc.moveDown(0.8);

    // Toplamlar (sağa dayalı küçük tablo)
    const totals: [string, string][] = [
      ["Ara Toplam", `₺ ${fmt(order.gross)}`],
      [`İskonto (%${order.discountPct})`, `₺ ${fmt(order.discount)}`],
      ["KDV", order.vatApplied ? `%20 — ₺ ${fmt(order.vatAmount)}` : "Uygulanmadı"],
      ["GENEL TOPLAM", `₺ ${fmt(order.net)}`],
    ];
    const tx = 50 + pageWidth * 0.5;
    const tw = pageWidth * 0.5;
    totals.forEach(([k, v], i) => {
      const bold = i === totals.length - 1;
      const y = doc.y;
      if (bold) doc.rect(tx, y - 2, tw, 18).fill("#f5f0e6");
      doc.fillColor(bold ? BRAND : "#000").font(bold ? FONT_BOLD : FONT).fontSize(10);
      doc.text(k, tx + 4, y, { width: tw * 0.55 });
      doc.text(v, tx + tw * 0.55, y, { width: tw * 0.45 - 6, align: "right" });
      doc.y = y + 18;
      doc.x = 50;
    });

    // Alt bilgi — sayfa alt marjının üstünde tek satır (taşarsa pdfkit yeni
    // sayfa açacağı için lineBreak kapalı)
    doc
      .font(FONT)
      .fontSize(8)
      .fillColor("#888")
      .text(
        "Olga Çerçeve — Sipariş Hattı: 0850 305 75 45 · olgacerceve.com",
        50,
        doc.page.height - 68,
        { width: pageWidth, align: "center", lineBreak: false }
      );

    doc.end();
  });
}
