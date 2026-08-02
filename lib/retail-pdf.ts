// Perakende sipariş üretim PDF'i — "OLGA ÇERÇEVE Sipariş PDF Üretim Sistemi"
// talimatının pdfkit uyarlaması. Tek sayfa A4: koyu altın bantlar, ölçekli
// üretim diyagramı (çerçeve/paspartu/eser/cam/+2mm pay), ölçü okları,
// turuncu uyarı bandı, göstergeler ve askı yönü.
// Kurallar: ölçüler .1f virgüllü; e-posta ASLA yazılmaz; cam yoksa overlay çizilmez.

import path from "path";
import PDFDocument from "pdfkit";
import type { SavedRetailOrder, RetailItem } from "./retail-orders";

const FONT = path.join(process.cwd(), "assets", "DejaVuSans.ttf");
const FONT_BOLD = path.join(process.cwd(), "assets", "DejaVuSans-Bold.ttf");

// Renk paleti (talimat §3)
const DARK_GOLD = "#5C4500";
const GOLD = "#8B6914";
const FRAME_FILL = "#5C3D11";
const ART_FILL = "#D8D8D8";
const MAT_PLAIN = "#F0EAD6";
const MAT_VELVET = "#4A235A";
const MAT_GOLDSILVER = "#D4C9A8";
const MAT_INNER = "#D9CDB8";
const PAY_COL = "#E8C000";
const SHADOW = "#AAAAAA";
const RED = "#C0392B";
const BLUE = "#1A5276";
const ORANGE = "#E67E22";
const PURPLE = "#7D3C98";
const GREEN = "#27AE60";
const DARK_GRAY = "#333333";
const MID_GRAY = "#DDDDDD";
const LABEL_GOLD = "#D4B483";
const CREAM = "#F5EDD6";

const MM = 2.83465; // 1 mm = 2.83465 pt
const mm = (v: number) => v * MM;
const PAY = 2; // +2mm kesim payı (standart)

const cm1 = (mmVal: number) => (mmVal / 10).toFixed(1).replace(".", ",");

const fmtTL = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }) + " TL";

interface ItemGeom {
  artW: number; // mm
  artH: number;
  frameW: number; // eser + paspartu kenarları
  frameH: number;
  hasMat: boolean;
  aski: "Yatay" | "Dikey" | "Kare";
}

function geom(it: RetailItem): ItemGeom {
  const artW = it.artWidthUnit === "cm" ? it.artWidth * 10 : it.artWidth;
  const artH = it.artHeightUnit === "cm" ? it.artHeight * 10 : it.artHeight;
  const hasMat =
    it.matType !== "Paspartu Yok" &&
    (it.matTop > 0 || it.matBottom > 0 || it.matLeft > 0 || it.matRight > 0);
  const frameW = artW + (hasMat ? it.matLeft + it.matRight : 0);
  const frameH = artH + (hasMat ? it.matTop + it.matBottom : 0);
  const aski = frameW > frameH ? "Yatay" : frameH > frameW ? "Dikey" : "Kare";
  return { artW, artH, frameW, frameH, hasMat, aski };
}

// Koyu zeminde etiketler beyaz yazılır
function isDark(hex: string): boolean {
  const m = /^#([0-9a-f]{6})$/i.exec(hex || "");
  if (!m) return false;
  const v = parseInt(m[1], 16);
  const r = (v >> 16) & 255, g = (v >> 8) & 255, b = v & 255;
  return 0.299 * r + 0.587 * g + 0.114 * b < 140;
}

function matColor(it: RetailItem, inner = false): string {
  const hex = inner ? it.innerMatColorHex : it.matColorHex;
  if (hex && hex.startsWith("#")) return hex;
  const t = inner ? it.innerMatType : it.matType;
  if (t.includes("Kadife") || t.includes("Premium")) return MAT_VELVET;
  if (t.includes("Altın")) return MAT_GOLDSILVER;
  if (inner) return MAT_INNER;
  return MAT_PLAIN;
}

export function generateRetailPdf(o: SavedRetailOrder): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margin: 0, font: FONT });
    const chunks: Buffer[] = [];
    doc.on("data", (c: Buffer) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
    doc.page.margins = { top: 0, bottom: 0, left: 0, right: 0 };

    const W = doc.page.width;
    const H = doc.page.height;
    // Tek sayfa: en fazla 4 diyagram. Kalemi olmayan (bozuk/eski) kayıtlarda
    // diyagram bölümü atlanır, PDF yine de üretilir.
    const items = (o.items || []).slice(0, 4);
    const single = items.length === 1;

    // ============ 1. HEADER (koyu altın bant, 24mm) ============
    doc.rect(0, 0, W, mm(24)).fill(DARK_GOLD);
    doc.font(FONT_BOLD).fontSize(20).fillColor("white");
    doc.text("OLGA Çerçeve", mm(15), mm(7));
    doc.font(FONT_BOLD).fontSize(9);
    doc.text(`Toplam: ${fmtTL(o.total)}`, W - mm(95), mm(6), {
      width: mm(80),
      align: "right",
    });
    const dateStr = new Date(o.createdAt).toLocaleDateString("tr-TR");
    const skuTxt = single ? items[0].frameCode : `${items.length} model`;
    doc.font(FONT).fontSize(9);
    doc.text(`${o.orderId}  |  SKU: ${skuTxt}  |  ${dateStr}`, W - mm(115), mm(14), {
      width: mm(100),
      align: "right",
    });

    // ============ 2. MÜŞTERİ (koyu altın bant, 26mm) ============
    const my = mm(26);
    doc.rect(0, my, W, mm(26)).fill(DARK_GOLD);
    doc.font(FONT_BOLD).fontSize(10).fillColor(CREAM);
    doc.text("MÜŞTERİ BİLGİLERİ", mm(15), my + mm(4));
    doc.font(FONT).fontSize(8.5).fillColor(LABEL_GOLD);
    doc.text("Ad Soyad:", mm(15), my + mm(12));
    doc.font(FONT_BOLD).fillColor("white");
    doc.text(o.customerName || "-", mm(40), my + mm(12));
    doc.font(FONT).fillColor(LABEL_GOLD);
    doc.text("Telefon:", W / 2, my + mm(12));
    doc.font(FONT_BOLD).fillColor("white");
    doc.text(o.customerPhone || "-", W / 2 + mm(18), my + mm(12));
    doc.font(FONT).fillColor(LABEL_GOLD);
    doc.text("Adres:", mm(15), my + mm(19));
    doc.font(FONT).fontSize(8).fillColor("white");
    doc.text(o.customerAddress || "-", mm(30), my + mm(19), { width: W - mm(45), height: mm(6), ellipsis: true });
    // NOT: e-posta bilerek yazılmıyor (talimat §2)

    // ============ 3. ÖZEL NOT BANTLARI ============
    let y = mm(52) + mm(3);
    const band = (color: string, text: string) => {
      doc.rect(mm(15), y, W - mm(30), mm(8)).fill(color);
      doc.font(FONT_BOLD).fontSize(9).fillColor("white");
      doc.text(text, mm(18), y + mm(2.2), { width: W - mm(36), height: mm(5), ellipsis: true });
      y += mm(10);
    };
    const printItems = items.filter((it) => it.printType && it.printType !== "Baskı Yok");
    if (printItems.length > 0) {
      band(PURPLE, `BASKI YAPILACAK — ${printItems.map((it) => it.printType).join(" | ")}`);
    }
    if (o.notes) {
      const low = o.notes.toLowerCase();
      const critical = low.includes("kırık") || low.includes("birleştir") || low.includes("kirik");
      band(critical ? RED : GREEN, o.notes.toUpperCase());
    }

    // ============ 4. ÜRETİM ÖNİZLEMESİ BAŞLIĞI ============
    y += mm(4);
    doc.font(FONT_BOLD).fontSize(11).fillColor(DARK_GOLD);
    doc.text("ÜRETİM ÖNİZLEMESİ", mm(15), y);
    doc.moveTo(mm(15), y + mm(6)).lineTo(W - mm(15), y + mm(6))
      .lineWidth(1.5).strokeColor(GOLD).stroke();

    // ============ 5. DİYAGRAM(LAR) ============
    const diagTop = y + mm(10);
    const diagMaxMM = single ? 62 : items.length === 2 ? 48 : items.length === 3 ? 38 : 34;
    const centers: [number, number][] =
      single
        ? [[W / 2, 0]]
        : items.length === 2
          ? [[W * 0.28, 0], [W * 0.72, 0]]
          : items.length === 3
            ? [[W * 0.18, 0], [W * 0.5, 0], [W * 0.82, 0]]
            : [[W * 0.3, 0], [W * 0.7, 0], [W * 0.3, 1], [W * 0.7, 1]];

    const rowH = mm(diagMaxMM + (single ? 26 : 18));
    let diagBottom = diagTop;

    items.forEach((it, idx) => {
      const g = geom(it);
      const [cxC, row] = centers[idx];
      const areaTop = diagTop + row * rowH + (single ? mm(8) : mm(6));
      const scale = Math.min(mm(diagMaxMM) / g.frameW, mm(diagMaxMM) / g.frameH);
      const cw = g.frameW * scale;
      const ch = g.frameH * scale;
      const cx = cxC - cw / 2;
      const cy = areaTop + (mm(diagMaxMM) - ch) / 2 + (single ? mm(4) : mm(2));
      const ft = (single ? 10 : 8) * scale * 10; // çerçeve profil kalınlığı görseli

      // Model etiketi (tek üründe çizilmez — SKU başlıkta ve göstergelerde var,
      // üstteki kırmızı ölçü okuyla çakışmasın)
      if (!single) {
        doc.font(FONT_BOLD).fontSize(8).fillColor(DARK_GRAY);
        doc.text(`#${idx + 1}  ${it.frameCode}`, cx - mm(10), areaTop - mm(4), {
          width: cw + mm(20),
          align: "center",
        });
      }

      // Gölge → çerçeve
      doc.rect(cx + 4, cy + 4, cw, ch).fill(SHADOW);
      doc.rect(cx, cy, cw, ch).fillColor(FRAME_FILL).fill();
      doc.rect(cx, cy, cw, ch).lineWidth(single ? 3 : 2).strokeColor("black").stroke();

      // Paspartu katmanları
      let ix = cx + ft, iy = cy + ft, iw = cw - ft * 2, ih = ch - ft * 2;
      if (g.hasMat) {
        doc.rect(ix, iy, iw, ih).fill(matColor(it));
        // Kenar mm etiketleri (yalnız tek diyagramda)
        if (single) {
          const edgeCol = isDark(matColor(it)) ? "#FFFFFF" : DARK_GOLD;
          doc.font(FONT_BOLD).fontSize(6.5).fillColor(edgeCol);
          if (it.matTop > 0)
            doc.text(`Üst: ${it.matTop}mm`, ix, iy + it.matTop * scale * 0.35, { width: iw, align: "center" });
          if (it.matBottom > 0)
            doc.text(`Alt: ${it.matBottom}mm`, ix, iy + ih - it.matBottom * scale * 0.55, { width: iw, align: "center" });
          if (it.matLeft > 0) {
            doc.save();
            doc.rotate(-90, { origin: [ix + it.matLeft * scale * 0.5, iy + ih / 2] });
            doc.text(`Sol: ${it.matLeft}mm`, ix + it.matLeft * scale * 0.5 - mm(12), iy + ih / 2 - mm(1.5), { width: mm(24), align: "center" });
            doc.restore();
          }
          if (it.matRight > 0) {
            doc.save();
            doc.rotate(90, { origin: [ix + iw - it.matRight * scale * 0.5, iy + ih / 2] });
            doc.text(`Sağ: ${it.matRight}mm`, ix + iw - it.matRight * scale * 0.5 - mm(12), iy + ih / 2 - mm(1.5), { width: mm(24), align: "center" });
            doc.restore();
          }
        }
        ix += it.matLeft * scale;
        iy += it.matTop * scale;
        iw -= (it.matLeft + it.matRight) * scale;
        ih -= (it.matTop + it.matBottom) * scale;
        // Çift paspartu iç şeridi
        if (it.doubleMat) {
          const strip = Math.max(3, 5 * scale * 2);
          doc.rect(ix, iy, iw, ih).fill(matColor(it, true));
          ix += strip; iy += strip; iw -= strip * 2; ih -= strip * 2;
        }
      }

      // Sanat eseri + X çaprazları
      doc.rect(ix, iy, iw, ih).fillColor(ART_FILL).fill();
      doc.rect(ix, iy, iw, ih).lineWidth(0.5).strokeColor("#BBBBBB").stroke();
      doc.moveTo(ix, iy).lineTo(ix + iw, iy + ih).stroke();
      doc.moveTo(ix + iw, iy).lineTo(ix, iy + ih).stroke();

      // Cam overlay — yalnız camlıysa
      if (it.glassType && it.glassType !== "Cam Yok") {
        doc.save();
        doc.opacity(0.22);
        doc.rect(ix, iy, iw, ih).fill("#BDE3F5");
        doc.restore();
      }

      // +2mm pay çizgisi (sarı)
      const ps = PAY * scale * 5;
      doc.rect(ix - ps / 2, iy - ps / 2, iw + ps, ih + ps)
        .lineWidth(2.5).strokeColor(PAY_COL).stroke();

      // İç ölçü yazısı — alana sığmazsa küçült, yine sığmazsa hiç yazma
      // (çoklu düzende ölçüler zaten diyagram altında yazılı)
      const dimTxt = `${cm1(g.artW)} × ${cm1(g.artH)} cm`;
      let dimFs = single ? 7 : 6;
      doc.font(FONT_BOLD).fontSize(dimFs);
      if (doc.widthOfString(dimTxt) > iw - 4) {
        dimFs -= 1.5;
        doc.fontSize(dimFs);
      }
      if (doc.widthOfString(dimTxt) <= iw - 2) {
        doc.fillColor(DARK_GRAY);
        doc.text(dimTxt, ix, iy + ih / 2 - mm(2.5), { width: iw, align: "center" });
        doc.font(FONT).fontSize(Math.max(4.5, dimFs - 0.5)).fillColor(GOLD);
        doc.text(`+${PAY}mm paylı`, ix, iy + ih / 2 + mm(0.5), { width: iw, align: "center" });
      }

      if (single) {
        // Ölçü okları — üst kırmızı (çerçeve dış ölçüsü)
        const topY = cy - mm(4);
        doc.lineWidth(1.2).strokeColor(RED);
        doc.moveTo(cx, topY).lineTo(cx + cw, topY).stroke();
        doc.moveTo(cx, topY - 4).lineTo(cx, topY + 4).stroke();
        doc.moveTo(cx + cw, topY - 4).lineTo(cx + cw, topY + 4).stroke();
        doc.font(FONT_BOLD).fontSize(9).fillColor(RED);
        doc.text(
          `${cm1(g.frameW)} × ${cm1(g.frameH)} cm  (+${PAY}mm pay)`,
          cx - mm(20), topY - mm(5.5),
          { width: cw + mm(40), align: "center" }
        );
        // Sağ dikey kırmızı ok
        const rx = cx + cw + mm(6);
        doc.strokeColor(RED);
        doc.moveTo(rx, cy).lineTo(rx, cy + ch).stroke();
        doc.moveTo(rx - 4, cy).lineTo(rx + 4, cy).stroke();
        doc.moveTo(rx - 4, cy + ch).lineTo(rx + 4, cy + ch).stroke();
        doc.save();
        doc.rotate(90, { origin: [rx + mm(5), cy + ch / 2] });
        doc.font(FONT_BOLD).fontSize(8).fillColor(RED);
        doc.text(`${cm1(g.frameH)} cm`, rx + mm(5) - mm(20), cy + ch / 2 - mm(1.6), { width: mm(40), align: "center" });
        doc.restore();
        // Alt mavi ok — sanat eseri
        const botY = cy + ch + mm(5);
        doc.lineWidth(1.2).strokeColor(BLUE);
        doc.moveTo(ix, botY).lineTo(ix + iw, botY).stroke();
        doc.moveTo(ix, botY - 3.5).lineTo(ix, botY + 3.5).stroke();
        doc.moveTo(ix + iw, botY - 3.5).lineTo(ix + iw, botY + 3.5).stroke();
        doc.font(FONT_BOLD).fontSize(9).fillColor(BLUE);
        doc.text(`Sanat eseri: ${cm1(g.artW)} × ${cm1(g.artH)} cm`, cx - mm(20), botY + mm(1.5), { width: cw + mm(40), align: "center" });
        diagBottom = Math.max(diagBottom, botY + mm(8));
      } else {
        // Çoklu düzen: altta kompakt ölçü satırı
        const botY = areaTop + mm(diagMaxMM) + mm(4);
        doc.font(FONT).fontSize(7).fillColor(DARK_GRAY);
        doc.text(
          `Çerçeve: ${cm1(g.frameW)}×${cm1(g.frameH)} cm  ·  Eser: ${cm1(g.artW)}×${cm1(g.artH)} cm`,
          cxC - mm(45), botY, { width: mm(90), align: "center" }
        );
        diagBottom = Math.max(diagBottom, botY + mm(6));
      }
    });

    // Çoklu düzende dikey ayırıcı
    if (items.length === 2) {
      doc.moveTo(W / 2, diagTop).lineTo(W / 2, diagBottom - mm(4))
        .lineWidth(0.5).strokeColor(MID_GRAY).stroke();
    }

    // ============ 6. TURUNCU UYARI BANDI ============
    y = diagBottom + mm(3);
    const warnLines: string[] = [];
    items.forEach((it, idx) => {
      const g = geom(it);
      const pre = items.length > 1 ? `#${idx + 1} ` : "";
      if (it.glassType && it.glassType !== "Cam Yok") {
        warnLines.push(`${pre}ÖZEL CAM: ${it.glassType}`);
      }
      if (g.hasMat) {
        const renk = it.matColor && it.matColor !== "-" ? ` ${it.matColor}` : "";
        warnLines.push(
          `${pre}PASPARTU: ${it.matType}${renk} — Ü${it.matTop}/S${it.matRight}/A${it.matBottom}/S${it.matLeft}mm` +
          (it.doubleMat ? ` + İç: ${it.innerMatType} ${it.innerMatColor}` : "")
        );
        warnLines.push(
          `${pre}Sanat eseri ${cm1(g.artW)}×${cm1(g.artH)} cm → Paspartu eklenince çerçeve ${cm1(g.frameW)}×${cm1(g.frameH)} cm olur`
        );
      }
    });
    if (warnLines.length > 0) {
      const bandH = mm(4) + warnLines.length * mm(4.6) + mm(2);
      doc.rect(mm(15), y, W - mm(30), bandH).fill(ORANGE);
      doc.font(FONT_BOLD).fontSize(8.5).fillColor("white");
      warnLines.forEach((l, i) => {
        doc.text(`⚠  ${l}`, mm(19), y + mm(2.5) + i * mm(4.6), { width: W - mm(38), height: mm(5), ellipsis: true });
      });
      y += bandH + mm(3);
    }

    // ============ 7. AYIRICI + GÖSTERGELER ============
    doc.moveTo(mm(15), y).lineTo(W - mm(15), y).lineWidth(1.5).strokeColor(GOLD).stroke();
    y += mm(4);
    doc.font(FONT_BOLD).fontSize(11).fillColor(DARK_GOLD);
    doc.text("GÖSTERGELER", mm(15), y);
    y += mm(7);

    type LegendRow = { box?: [string, string]; label: string; value: string };
    const rows: LegendRow[] = [];
    items.forEach((it, idx) => {
      const pre = items.length > 1 ? `#${idx + 1} ` : "";
      rows.push({ box: [FRAME_FILL, "black"], label: `${pre}Çerçeve profili  — `, value: it.frameCode });
      const g = geom(it);
      if (g.hasMat) {
        rows.push({
          box: [matColor(it), "#B8A88A"],
          label: `${pre}${it.doubleMat ? "Dış " : ""}Paspartu  — `,
          value: `${it.matType}${it.matColor !== "-" ? ` (${it.matColor})` : ""}`,
        });
        if (it.doubleMat) {
          rows.push({
            box: [matColor(it, true), "#B8A88A"],
            label: `${pre}İç Paspartu  — `,
            value: `${it.innerMatType}${it.innerMatColor !== "-" ? ` (${it.innerMatColor})` : ""} · montaj ${it.altMontaj}mm`,
          });
        }
      }
      if (it.glassType && it.glassType !== "Cam Yok") {
        rows.push({ box: ["#BDE3F5", "#5DADE2"], label: `${pre}Cam  — `, value: it.glassType });
      }
    });
    const g0 = items.length > 0 ? geom(items[0]) : null;
    rows.push({
      box: [ART_FILL, "#BBBBBB"],
      label: "Sanat eseri  — ",
      value:
        single && g0
          ? `${cm1(g0.artW)} × ${cm1(g0.artH)} cm  (+${PAY}mm pay)`
          : `${items.length} eser (+${PAY}mm pay)`,
    });
    rows.push({ box: [PAY_COL, "#B8860B"], label: `+${PAY} mm pay  — `, value: "Sarı çizgi ile gösterilmiştir" });
    if (items.every((it) => !it.glassType || it.glassType === "Cam Yok")) {
      rows.push({ label: "Cam  — ", value: "Yok" });
    }
    if (items.every((it) => !geom(it).hasMat)) {
      rows.push({ label: "Paspartu  — ", value: "Yok" });
    }
    const askiTxt =
      single && g0
        ? g0.aski
        : items.map((it, i) => `#${i + 1} ${geom(it).aski}`).join(", ") || "-";
    rows.push({ label: "Askı  — ", value: askiTxt });

    // Satır yüksekliği: footer'a sığacak şekilde daralt
    const footerTop = H - mm(22);
    const avail = footerTop - y - mm(2);
    const rh = Math.max(mm(4.6), Math.min(mm(7.5), avail / rows.length));
    const fs = rh < mm(5.4) ? 8 : 10;

    rows.forEach((r, i) => {
      const by = y + i * rh;
      let tx = mm(24);
      if (r.box) {
        doc.rect(mm(15), by, mm(5), mm(5)).fillColor(r.box[0]).fill();
        doc.rect(mm(15), by, mm(5), mm(5)).lineWidth(1).strokeColor(r.box[1]).stroke();
      } else {
        tx = mm(24);
      }
      doc.font(FONT).fontSize(fs).fillColor(DARK_GRAY);
      doc.text(r.label, tx, by + mm(0.8), { continued: true });
      doc.font(FONT_BOLD).text(r.value);
    });

    // Askı yön oku (tek üründe)
    if (single && g0) {
      const lastY = y + (rows.length - 1) * rh + mm(2.5);
      doc.font(FONT_BOLD).fontSize(fs);
      const sx = mm(24) + doc.widthOfString(`Askı  — ${askiTxt}`) + mm(10);
      doc.lineWidth(1.5).strokeColor(DARK_GRAY);
      if (g0.aski === "Yatay") {
        doc.moveTo(sx - 5, lastY).lineTo(sx + 5, lastY).stroke();
        doc.moveTo(sx + 1.5, lastY + 3.5).lineTo(sx + 5, lastY).stroke();
        doc.moveTo(sx + 1.5, lastY - 3.5).lineTo(sx + 5, lastY).stroke();
        doc.moveTo(sx - 1.5, lastY + 3.5).lineTo(sx - 5, lastY).stroke();
        doc.moveTo(sx - 1.5, lastY - 3.5).lineTo(sx - 5, lastY).stroke();
      } else if (g0.aski === "Dikey") {
        doc.moveTo(sx, lastY + 5).lineTo(sx, lastY - 5).stroke();
        doc.moveTo(sx - 3.5, lastY + 1.5).lineTo(sx, lastY + 5).stroke();
        doc.moveTo(sx + 3.5, lastY + 1.5).lineTo(sx, lastY + 5).stroke();
        doc.moveTo(sx - 3.5, lastY - 1.5).lineTo(sx, lastY - 5).stroke();
        doc.moveTo(sx + 3.5, lastY - 1.5).lineTo(sx, lastY - 5).stroke();
      }
    }

    // ============ 8. FOOTER ============
    doc.moveTo(mm(15), H - mm(18)).lineTo(W - mm(15), H - mm(18))
      .lineWidth(0.5).strokeColor(MID_GRAY).stroke();
    doc.font(FONT_BOLD).fontSize(8).fillColor(GOLD);
    doc.text("OLGA Çerçeve", mm(15), H - mm(13));
    doc.font(FONT).fontSize(8).fillColor("#888888");
    doc.text("0850 305 75 45", 0, H - mm(13), { width: W, align: "center" });
    doc.text("www.olgacerceve.com", W - mm(75), H - mm(13), { width: mm(60), align: "right" });

    doc.end();
  });
}
