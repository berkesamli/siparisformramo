function myFunction() {
  
}
// ==============================
// File: Code.gs  (Google Apps Script)
// ==============================

const EMAIL_TO        = "olgacercevee@gmail.com";
const ORDERS_SHEET    = "Siparişler";
const ITEMS_SHEET     = "Sipariş Kalemleri";
const DAILY_SHEET     = "Günlük";
const CATALOG_SHEET   = "çerçeve bilgiler";
const TIMEZONE        = "Europe/Istanbul";

const SPREADSHEET_ID  = "113SwuC5QNJBvVjlQ247mk_GzjegMKmb7MLaTn-9v_2I";

const ORDER_PREFIX    = "OLG";
const ORDER_PAD       = 5;
const PDF_FOLDER_NAME = "Sipariş PDF";
const BRAND_COLOR     = "#8b4b00";
const HEADER_BG       = "#f7f3ef";

/** Spreadsheet erişimi: önce getActive, sonra openById dener */
function getSS_(){
  try { const ss = SpreadsheetApp.getActiveSpreadsheet(); if(ss) return ss; } catch(e){}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function hasUi_(){ try{ SpreadsheetApp.getUi(); return true; }catch(e){ return false; } }

/** Apps Script editöründen çalıştırıp yetkilendirme izni almak için */
function testErisim(){
  const ss = getSS_();
  Logger.log("Spreadsheet adı: " + ss.getName());
  Logger.log("Sheet sayısı: " + ss.getSheets().length);
  ss.getSheets().forEach(s => Logger.log("  - " + s.getName()));
}

function onOpen(){
  if(!hasUi_()) return;
  SpreadsheetApp.getUi().createMenu("🧰 Sipariş Sistemi")
    .addItem("Formu Aç", "openForm_")
    .addSeparator()
    .addItem("Sayaç Sıfırla", "resetOrderCounter_")
    .addToUi();
}
function openForm_(){
  if(!hasUi_()) return;
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if(!url){
    ui.alert("Web App yayımlı değil. Deploy → New deployment ile yayımlayın.");
    return;
  }
  ui.showModalDialog(
    HtmlService.createHtmlOutput(`<p>Form URL:</p><p><a target="_blank" href="${url}">${url}</a></p>`).setWidth(420).setHeight(120),
    "Sipariş Formu"
  );
}
function resetOrderCounter_(){
  PropertiesService.getScriptProperties().deleteProperty('ORDER_SEQ');
  if(hasUi_()) SpreadsheetApp.getUi().alert("Sipariş numarası sayacı sıfırlandı.");
}

function doGet(){
  try{
    return HtmlService.createHtmlOutputFromFile("Form").setTitle("Olga Çerçeve — Sipariş Formu");
  }catch(err){
    return ContentService.createTextOutput("Form.html bulunamadı.\n"+String(err)).setMimeType(ContentService.MimeType.TEXT);
  }
}
function submitFromHtml(data){ return processOrder_(data); }

function getCatalog(){
  const ss = getSS_();
  const sh = ss.getSheetByName(CATALOG_SHEET);
  if(!sh) throw new Error('Katalog sekmesi yok: ' + CATALOG_SHEET);
  const last = sh.getLastRow(); if(last < 2) return {};
  const values = sh.getRange(2,1,last-1,6).getValues();
  const map = {};
  values.forEach(r=>{
    const base = normalizeBaseCode_(r[0]); if(!base) return;
    const type = String(r[1]||"").trim().toLowerCase(); // metrelik|adetli
    const koliAdet  = numFromCell_(r[2]);
    const koliMetre = numFromCell_(r[3]);
    const priceUSD  = numFromCell_(r[4]);
    const boyLength = (type==="metrelik" && koliAdet>0 && koliMetre>0) ? (koliMetre/koliAdet) : 0;
    if(!(base in map)) map[base] = { base, type, koliAdet, koliMetre, priceUSD, boyLength };
  });
  return map;
}

function processOrder_(data){
  try{
    const payload = (data && typeof data==='object') ? data : {};
    const {orders, items, daily} = ensureSheets_();
    const catalog = getCatalog();
    if(!Array.isArray(payload.rows) || !payload.rows.length) throw new Error("Satır verisi boş.");

    const ts      = new Date();
    const dayKey  = Utilities.formatDate(ts, TIMEZONE, "dd.MM.yyyy");
    const orderId = getNextOrderId_();

    const rate       = Math.max(0, Number(payload.rate||0));
    const euroRate   = Math.max(0, Number(payload.euroRate||0));
    const vatApplied = !!payload.vatApplied;

    let lineNo=0, gross=0;
    const itemRows = [];

    (payload.rows||[]).forEach(row=>{
      if(!row) return;
      const kind = String(row.kind||"other").toLowerCase();

      if(kind==="frame"){
        const fullText = String(row.fullCode||"").trim();
        const base = normalizeBaseCode_(fullText);
        const c = catalog[base] || null;

        const unitSel = String(row.unit||"metre").toLowerCase(); // metre|boy|koli
        const qty     = Number(row.qty||0); if(qty<=0 || !fullText) return;

        const priceUSD    = Number(row.usd || (c?.priceUSD) || 0);
        const unitPriceTL = round2(priceUSD * rate); // yoksa 0

        let metres = 0;
        if(c?.type==="metrelik"){
          if(unitSel==="metre") metres = qty;
          else if(unitSel==="boy")  metres = qty * (c.boyLength||0);
          else if(unitSel==="koli") metres = qty * (c.koliMetre||0);
        }else{
          metres = qty; // fallback
        }
        metres = round2(metres);

        // Boy hesaplama: metre / boyLength
        let unitText = `${fmt(metres)} mt`;
        if(c?.boyLength > 0 && metres > 0){
          const boyCount = Math.round(metres / c.boyLength * 10) / 10;
          unitText = `${fmt(metres)} mt (${boyCount} boy)`;
        }
        const lineTotal = round2(metres * unitPriceTL);

        lineNo++; gross = round2(gross + lineTotal);
        itemRows.push([orderId, lineNo, fullText || base || "Ürün", unitText, unitPriceTL, lineTotal]);

      }else if(kind==="glass"){
        // Cam işleme - plaka bazlı
        const glassName = String(row.name||"Cam").trim();
        const sizeLabel = String(row.sizeLabel||"").trim();
        const plakaAdet = Number(row.plakaAdet||0);
        const m2PerPlaka = round2(Number(row.m2PerPlaka||0));
        const m2 = round2(Number(row.m2||0));
        const m2Price = round2(Number(row.m2Price||0));

        if(m2 <= 0 || plakaAdet <= 0) return;

        const unitText = `${plakaAdet} plaka × ${fmt(m2PerPlaka)} m² = ${fmt(m2)} m² (${sizeLabel})`;
        const lineTotal = round2(m2 * m2Price);

        lineNo++; gross = round2(gross + lineTotal);
        itemRows.push([orderId, lineNo, glassName, unitText, m2Price, lineTotal]);

      }else if(kind==="ayna"){
        // Ayna işleme - plaka bazlı (cam ile aynı mantık)
        const aynaName = String(row.name||"Ayna").trim();
        const sizeLabel = String(row.sizeLabel||"").trim();
        const plakaAdet = Number(row.plakaAdet||0);
        const m2PerPlaka = round2(Number(row.m2PerPlaka||0));
        const m2 = round2(Number(row.m2||0));
        const m2Price = round2(Number(row.m2Price||0));

        if(m2 <= 0 || plakaAdet <= 0) return;

        const unitText = `${plakaAdet} plaka × ${fmt(m2PerPlaka)} m² = ${fmt(m2)} m² (${sizeLabel})`;
        const lineTotal = round2(m2 * m2Price);

        lineNo++; gross = round2(gross + lineTotal);
        itemRows.push([orderId, lineNo, aynaName, unitText, m2Price, lineTotal]);

      }else if(kind==="technical"){
        // Teknik malzeme işleme - kutu bazlı
        const productCode = String(row.code||"").trim();
        const productName = String(row.name||"").trim();
        const category = String(row.category||"Teknik Malzeme").trim();
        const kartonKodu = String(row.kartonKodu||"").trim();
        const kutuAdet = Number(row.kutuAdet||0);
        const adetPerKutu = Number(row.adetPerKutu||0);
        const totalAdet = Number(row.totalAdet||0);
        const priceEUR = round2(Number(row.priceEUR||0));
        const priceTL = round2(Number(row.priceTL||0));
        const euroRate = Number(row.euroRate||0);
        const kutuPriceTL = round2(Number(row.kutuPriceTL||0));

        if(kutuAdet <= 0) return;

        // Karton kodu varsa ürün adına ekle
        const fullName = kartonKodu
          ? `${productName} (${kartonKodu})`
          : productName;
        // TL veya EUR fiyat gösterimi
        const priceInfo = priceTL > 0 ? `₺${fmt(priceTL)}/adet` : `€${fmt(priceEUR)}/adet`;
        const unitText = `${kutuAdet} adet × ${adetPerKutu} = ${totalAdet} adt (${priceInfo})`;
        const lineTotal = round2(kutuAdet * kutuPriceTL);

        lineNo++; gross = round2(gross + lineTotal);
        itemRows.push([orderId, lineNo, fullName, unitText, kutuPriceTL, lineTotal]);

      }else{
        const name = String(row.name||"").trim();
        const qty  = Number(row.qty||0); if(qty<=0 || !name) return;
        const unitPriceTL = round2(Number(row.unitPriceTL||0)); // boşsa 0
        const unitText = `${qty} adt`;
        const lineTotal = round2(qty * unitPriceTL);

        lineNo++; gross = round2(gross + lineTotal);
        itemRows.push([orderId, lineNo, name || "Diğer", unitText, unitPriceTL, lineTotal]);
      }
    });
    if(!itemRows.length) throw new Error("Geçerli satır yok.");

    const pct = Math.max(0, Number(payload.discountPct||0))/100;
    const discount = round2(gross * pct);
    const afterDiscount = round2(Math.max(0, gross - discount));
    const vatAmount = vatApplied ? round2(afterDiscount * 0.20) : 0;
    const net = round2(afterDiscount + vatAmount);

    if(itemRows.length){
      items.getRange(items.getLastRow()+1,1,itemRows.length,itemRows[0].length).setValues(itemRows);
    }
    orders.appendRow([
      ts, orderId, (payload.employee||""), (payload.customer||""), (payload.note||""),
      rate, gross, discount, vatApplied ? "KDV %20" : "KDV Yok", vatAmount, net
    ]);

    writeDailyHuman_(daily, ts, dayKey, orderId, (payload.customer||""), itemRows, {
      discountPct: Number(payload.discountPct||0), discount, gross, vatApplied, vatAmount, net
    });

    const pdfInfo = createOrderPdf_(orderId, ts, payload, itemRows, gross, discount, vatApplied, vatAmount, net, rate, euroRate);

    MailApp.sendEmail({
      to: EMAIL_TO,
      subject: `Yeni Sipariş ${orderId} — ${payload.customer||""} — ₺ ${fmt(net)}`,
      htmlBody: buildEmailBody_(orderId, ts, payload, itemRows, gross, discount, vatApplied, vatAmount, net, rate, euroRate),
      attachments: [pdfInfo.blob]
    });

    return {ok:true, orderId, gross, discount, vatApplied, vatAmount, net, pdfUrl: pdfInfo.url};
  }catch(err){
    return {ok:false, error:String(err)};
  }
}

function getNextOrderId_(){
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const props = PropertiesService.getScriptProperties();
    let n = Number(props.getProperty('ORDER_SEQ') || '0');
    n = isFinite(n) ? n + 1 : 1;
    props.setProperty('ORDER_SEQ', String(n));
    return `${ORDER_PREFIX}${String(n).padStart(ORDER_PAD,'0')}`;
  } finally { lock.releaseLock(); }
}

function buildEmailBody_(orderId, ts, data, itemRows, gross, discount, vatApplied, vatAmount, net, rate, euroRate){
  const note = `${esc(data.note||"-")}<br>Kurlar: ${fmt(rate)} TL/USD | ${fmt(euroRate||0)} TL/EUR`;
  return `
  <div style="font:14px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#222">
    <h3 style="margin:0 0 8px">Yeni sipariş</h3>
    <b>Sipariş No:</b> ${orderId}<br>
    <b>Çalışan:</b> ${esc(data.employee||"")}<br>
    <b>Müşteri:</b> ${esc(data.customer||"")}<br>
    <b>Not:</b> ${note}<br><br>
    <table border="1" cellpadding="6" cellspacing="0">
      <tr><th>#</th><th>Ürün</th><th>Birim</th><th>Birim Fiyat (₺)</th><th>Tutar (₺)</th></tr>
      ${itemRows.map(r=>`<tr>
        <td>${r[1]}</td><td>${esc(r[2])}</td><td>${esc(r[3])}</td>
        <td>₺ ${fmt(r[4])}</td><td>₺ ${fmt(r[5])}</td>
      </tr>`).join("")}
    </table><br>
    <b>Ara Toplam:</b> ₺ ${fmt(gross)}<br>
    <b>İndirim:</b> ${Number(data.discountPct||0).toFixed(2)}% — ₺ ${fmt(discount)}<br>
    <b>KDV:</b> ${vatApplied ? "%20 — ₺ "+fmt(vatAmount) : "Uygulanmadı"}<br>
    <b>Genel Toplam:</b> ₺ ${fmt(net)}<br>
    <div style="margin-top:8px;color:#666">Zaman: ${Utilities.formatDate(ts,TIMEZONE,"dd.MM.yyyy HH:mm")}</div>
  </div>`;
}

function ensureSheets_(){
  const ss = getSS_();
  let orders = ss.getSheetByName(ORDERS_SHEET);
  if(!orders){
    orders = ss.insertSheet(ORDERS_SHEET);
    orders.appendRow(["Tarih","Sipariş No","Çalışan","Müşteri","Not","Kur","Ara Toplam","İndirim (₺)","KDV","KDV Tutarı","Genel Toplam"]);
    orders.setFrozenRows(1);
    orders.getRange("A:A").setNumberFormat("dd.mm.yyyy hh:mm");
    orders.getRange("G:K").setNumberFormat("₺ #,##0.00");
  }
  let items = ss.getSheetByName(ITEMS_SHEET);
  if(!items){
    items = ss.insertSheet(ITEMS_SHEET);
    items.appendRow(["Sipariş No","Satır No","Ürün","Birim","Birim Fiyat (₺)","Satır Tutar (₺)"]);
    items.setFrozenRows(1);
    items.getRange("E:F").setNumberFormat("₺ #,##0.00");
  }
  let daily = ss.getSheetByName(DAILY_SHEET);
  if(!daily){
    daily = ss.insertSheet(DAILY_SHEET);
    daily.appendRow(["Tarih/Müşteri Başlığı","Sipariş No","Satır No","Ürün","Birim","Birim Fiyat","Satır Tutarı"]);
    daily.setFrozenRows(1);
  }
  return {orders, items, daily};
}

function writeDailyHuman_(sh, ts, dayKey, orderId, customer, itemRows, sums){
  const lastRow = sh.getLastRow();
  const marker  = `— ${dayKey} —`;
  const lastVal = lastRow > 1 ? sh.getRange(lastRow,1).getDisplayValue() : "";
  if(lastVal !== marker){
    const r = sh.getLastRow()+1;
    sh.getRange(r,1,1,7).merge()
      .setValue(marker).setFontWeight("bold").setHorizontalAlignment("center").setBackground(HEADER_BG);
  }
  const r2 = sh.getLastRow()+1;
  sh.getRange(r2,1,1,7).merge().setValue(customer || "Müşteri").setFontWeight("bold").setBackground("#fbf9f6");

  if(itemRows.length){
    const pretty = itemRows.map(r=>[r[0], r[1], r[2], r[3], `₺ ${fmt(r[4])} TL`, `₺ ${fmt(r[5])} TL`]);
    const start = sh.getLastRow()+1;
    sh.getRange(start,2,pretty.length,pretty[0].length).setValues(pretty);
  }

  const r3 = sh.getLastRow()+1;
  sh.getRange(r3,2,1,5).merge().setValue("Ara Toplam").setFontWeight("bold");
  sh.getRange(r3,7).setValue(`₺ ${fmt(sums.gross)} TL`);

  const r4 = sh.getLastRow()+1;
  sh.getRange(r4,2,1,5).merge().setValue(`İndirim (${(sums.discountPct||0).toFixed(2)}%)`).setFontWeight("bold");
  sh.getRange(r4,7).setValue(`₺ ${fmt(sums.discount)} TL`);

  const r5 = sh.getLastRow()+1;
  sh.getRange(r5,2,1,5).merge().setValue("KDV").setFontWeight("bold");
  sh.getRange(r5,7).setValue(sums.vatApplied ? `₺ ${fmt(sums.vatAmount)} TL` : "Uygulanmadı");

  const r6 = sh.getLastRow()+1;
  sh.getRange(r6,2,1,5).merge().setValue("Genel Toplam").setFontWeight("bold").setBackground("#fff9e8");
  sh.getRange(r6,7).setValue(`₺ ${fmt(sums.net)} TL`).setFontWeight("bold");
}

function createOrderPdf_(orderId, ts, data, itemRows, gross, discount, vatApplied, vatAmount, net, rate, euroRate){
  const folder = getOrCreateFolder_(PDF_FOLDER_NAME);
  const doc = DocumentApp.create(`Sipariş ${orderId}`);
  const b = doc.getBody(); b.clear();

  b.appendParagraph("OLGA ÇERÇEVE").setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setForegroundColor(BRAND_COLOR).setBold(true);
  b.appendParagraph("Sipariş Özeti").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  [
    `Sipariş No: ${orderId}`,
    `Tarih: ${Utilities.formatDate(ts, TIMEZONE, "dd.MM.yyyy HH:mm")}`,
    `Çalışan: ${data.employee||""}`,
    `Müşteri: ${data.customer||""}`,
    `Not: ${(data.note||"-")}`,
    `Kurlar: ${fmt(rate)} TL/USD | ${fmt(euroRate||0)} TL/EUR`
  ].forEach(t=>b.appendParagraph(t));
  b.appendParagraph(" ");

  const tbl = b.appendTable();
  const head = tbl.appendTableRow();
  ["#","Ürün","Birim","Birim Fiyat (₺)","Tutar (₺)"].forEach(h=>{
    head.appendTableCell(h).setBackgroundColor(HEADER_BG).setBold(true);
  });
  itemRows.forEach(r=>{
    const row = tbl.appendTableRow();
    row.appendTableCell(String(r[1]));
    row.appendTableCell(String(r[2]));
    row.appendTableCell(String(r[3]));
    row.appendTableCell("₺ " + fmt(r[4]));
    row.appendTableCell("₺ " + fmt(r[5]));
  });
  b.appendParagraph(" ");

  const sums = b.appendTable([
    ["Ara Toplam",   "₺ " + fmt(gross)],
    ["İndirim",      `${Number(data.discountPct||0).toFixed(2)}% — ₺ ${fmt(discount)}`],
    ["KDV",          vatApplied ? "%20 — ₺ " + fmt(vatAmount) : "Uygulanmadı"],
    ["Genel Toplam", "₺ " + fmt(net)]
  ]);
  for(let i=0;i<4;i++) sums.getRow(i).getCell(0).setBold(true).setForegroundColor(BRAND_COLOR);

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  const blob = file.getAs(MimeType.PDF).setName(`Siparis_${orderId}.pdf`);
  const pdf  = folder.createFile(blob);
  file.setTrashed(true);
  return {url: pdf.getUrl(), blob: pdf.getBlob()};
}

/* Helpers */
function normalizeBaseCode_(s){
  const str = (s||"").toString().toUpperCase().replace(/\s+/g," ").trim()
               .replace(/\bSERISI\b|\bSERİSİ\b/g,"").replace(/\s+/g,"");
  const i = str.indexOf("-"); return i>0 ? str.substring(0,i) : str;
}
function getOrCreateFolder_(name){
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function numFromCell_(v){
  if (v == null) return 0;
  const n = parseFloat(String(v).trim().replace(/\s+/g,'').replace(/[A-Za-z]+/g,'').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function round2(n){ return Math.round(Number(n||0)*100)/100; }
function fmt(n){ return (Number(n)||0).toLocaleString('tr-TR',{minimumFractionDigits:2, maximumFractionDigits:2}); }
function esc(s){ return String(s||"").replace(/[&<>"']/g, m=>({"&":"&amp;","<":"&lt;","&gt;":"&gt;","\"":"&quot;","'":"&#39;"}[m])); }
