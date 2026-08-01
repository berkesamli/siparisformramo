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
const LOGO = path.join(process.cwd(), "assets", "olga-logo.png");

const BRAND = "#a9660d";
const BRAND_DARK = "#7c4a06";
const CREAM = "#faf7f1";
const INK = "#26201a";
const GRAY = "#8a8072";
const LINE = "#e8e0d2";

const M = 46; // sayfa marjı

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export function generateOrderPdf(order: PdfOrder): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    // font seçeneği ile başlatınca pdfkit standart (Helvetica) fontlarını hiç
    // yüklemez — Türkçe karakterler ve Vercel paketi için gereklidir.
    const doc = new PDFDocument({ size: "A4", margin: M, font: FONT });
    const chunks: Buffer[] = [];
    doc.on("data", (c: Buffer) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    const W = doc.page.width;
    const pageWidth = W - M * 2;

    // ===== Üst alan: beyaz zemin + gerçek logo, altta bronz ayraç =====
    const bandH = 96;
    doc.image(LOGO, M, 24, { height: 42 });
    doc
      .font(FONT)
      .fontSize(7.5)
      .fillColor(BRAND)
      .text("ÇERÇEVE", M + 2, 72, { characterSpacing: 6.5 });

    doc
      .font(FONT_BOLD)
      .fontSize(16)
      .fillColor(BRAND)
      .text("SİPARİŞ FİŞİ", M, 28, { width: pageWidth, align: "right" });
    doc
      .font(FONT)
      .fontSize(9.5)
      .fillColor(GRAY)
      .text(order.orderId, M, 52, { width: pageWidth, align: "right" })
      .text(order.dateStr, M, 66, { width: pageWidth, align: "right" });

    doc.rect(0, bandH, W, 2.5).fill(BRAND);
    doc.rect(0, bandH + 2.5, W, 1.2).fill("#c99b4a");

    // ===== Bilgi kutuları =====
    let y = bandH + 24;
    const half = pageWidth / 2 - 8;

    const infoBox = (
      x: number,
      title: string,
      rows: [string, string][],
      w: number
    ) => {
      const rowH = 16;
      const h = 26 + rows.length * rowH;
      doc.roundedRect(x, y, w, h, 6).fill(CREAM);
      doc
        .font(FONT_BOLD)
        .fontSize(8)
        .fillColor(BRAND)
        .text(title.toUpperCase(), x + 12, y + 9, { characterSpacing: 1.5 });
      rows.forEach(([k, v], i) => {
        const ry = y + 26 + i * rowH;
        doc.font(FONT).fontSize(9).fillColor(GRAY).text(k, x + 12, ry, { width: 70 });
        doc
          .font(FONT_BOLD)
          .fontSize(9)
          .fillColor(INK)
          .text(v, x + 84, ry, { width: w - 96 });
      });
      return h;
    };

    const leftRows: [string, string][] = [["Müşteri", order.customer || "-"]];
    const rightRows: [string, string][] = [["Çalışan", order.employee]];
    if (order.status) rightRows.push(["Durum", order.status]);

    const h1 = infoBox(M, "Müşteri Bilgileri", leftRows, half);
    const h2 = infoBox(M + half + 16, "Sipariş Bilgileri", rightRows, half);
    y += Math.max(h1, h2) + 22;

    // ===== Ürün tablosu =====
    const cols = [26, pageWidth * 0.36, pageWidth * 0.29, pageWidth * 0.14, pageWidth * 0.16];
    const colX: number[] = [];
    {
      let x = M;
      for (const w of cols) {
        colX.push(x);
        x += w;
      }
    }

    const cellHeight = (texts: string[]) =>
      Math.max(
        ...texts.map((t, i) =>
          doc.heightOfString(t, { width: cols[i] - 12 })
        )
      ) + 11;

    const drawCells = (
      texts: string[],
      opts: { header?: boolean; zebra?: boolean } = {}
    ) => {
      doc.font(opts.header ? FONT_BOLD : FONT).fontSize(opts.header ? 8.5 : 9);
      const h = cellHeight(texts);
      if (y + h > doc.page.height - 150) {
        doc.addPage();
        y = M;
      }
      if (opts.header) {
        doc.roundedRect(M, y, pageWidth, h, 4).fill(BRAND);
      } else if (opts.zebra) {
        doc.rect(M, y, pageWidth, h).fill(CREAM);
      }
      texts.forEach((t, i) => {
        doc
          .font(opts.header ? FONT_BOLD : i === 1 ? FONT_BOLD : FONT)
          .fontSize(opts.header ? 8.5 : 9)
          .fillColor(opts.header ? "#ffffff" : INK)
          .text(t, colX[i] + 6, y + 6, {
            width: cols[i] - 12,
            align: i >= 3 ? "right" : "left",
            characterSpacing: opts.header ? 0.8 : 0,
          });
      });
      y += h;
      if (!opts.header) {
        doc
          .moveTo(M, y)
          .lineTo(M + pageWidth, y)
          .strokeColor(LINE)
          .lineWidth(0.5)
          .stroke();
      }
    };

    drawCells(["#", "ÜRÜN", "MİKTAR", "B. FİYAT", "TUTAR"], { header: true });
    order.lines.forEach((l, i) => {
      drawCells(
        [String(i + 1), l.name, l.unitText, `₺ ${fmt(l.unitPriceTL)}`, `₺ ${fmt(l.lineTotal)}`],
        { zebra: i % 2 === 1 }
      );
    });

    // ===== Toplamlar =====
    y += 16;
    const tw = pageWidth * 0.46;
    const tx = M + pageWidth - tw;
    const totals: [string, string, boolean][] = [
      ["Ara Toplam", `₺ ${fmt(order.gross)}`, false],
      [`İskonto (%${order.discountPct})`, `₺ ${fmt(order.discount)}`, false],
      ["KDV", order.vatApplied ? `%20 — ₺ ${fmt(order.vatAmount)}` : "Uygulanmadı", false],
      ["GENEL TOPLAM", `₺ ${fmt(order.net)}`, true],
    ];
    const totH = 20;
    if (y + totals.length * totH + 60 > doc.page.height - 80) {
      doc.addPage();
      y = M;
    }
    totals.forEach(([k, v, grand]) => {
      if (grand) {
        doc.roundedRect(tx, y - 1, tw, totH + 4, 5).fill(BRAND);
      }
      doc
        .font(grand ? FONT_BOLD : FONT)
        .fontSize(grand ? 11 : 9.5)
        .fillColor(grand ? "#ffffff" : GRAY)
        .text(k, tx + 10, y + 3, { width: tw * 0.55 });
      doc
        .font(FONT_BOLD)
        .fontSize(grand ? 11 : 9.5)
        .fillColor(grand ? "#ffffff" : INK)
        .text(v, tx + tw * 0.5, y + 3, { width: tw * 0.5 - 10, align: "right" });
      y += totH + (grand ? 6 : 2);
    });

    // ===== Not (Genel Toplam'ın altındaki boş alanda) =====
    if (order.note) {
      y += 10;
      doc.font(FONT).fontSize(9);
      const nh = doc.heightOfString(order.note, { width: pageWidth - 24 }) + 32;
      if (y + nh > doc.page.height - 90) {
        doc.addPage();
        y = M;
      }
      doc.roundedRect(M, y, pageWidth, nh, 6).fill(CREAM);
      doc
        .font(FONT_BOLD)
        .fontSize(8)
        .fillColor(BRAND)
        .text("NOT", M + 12, y + 9, { characterSpacing: 1.5 });
      doc
        .font(FONT)
        .fontSize(9)
        .fillColor(INK)
        .text(order.note, M + 12, y + 24, { width: pageWidth - 24 });
      y += nh;
    }

    // ===== Alt bilgi =====
    // Alt marjı sıfırla — yoksa pdfkit marj sınırını aşan footer için yeni sayfa açar
    doc.page.margins.bottom = 0;
    const fy = doc.page.height - 64;
    doc
      .moveTo(M, fy)
      .lineTo(M + pageWidth, fy)
      .strokeColor(BRAND)
      .lineWidth(0.8)
      .stroke();
    doc
      .font(FONT)
      .fontSize(7.5)
      .fillColor(GRAY)
      .text(
        "OLGA ÇERÇEVE  ·  Sipariş Hattı: 0850 305 75 45  ·  olgacerceve.com",
        M,
        fy + 10,
        { width: pageWidth, align: "center", lineBreak: false }
      );
    doc
      .fontSize(7)
      .fillColor("#b3a893")
      .text(
        "Ankara: Birlik Mah. 448. Cd. No:56 Çankaya — 0312 495 75 45   |   İstanbul: Masko Mobilya San. Sit. 3-B1 No:4 İkitelli — 0212 675 27 50",
        M,
        fy + 24,
        { width: pageWidth, align: "center", lineBreak: false }
      );

    doc.end();
  });
}
