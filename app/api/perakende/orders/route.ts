import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import {
  istanbulDateKey,
  lastNDateKeys,
  blobConfigured,
  PAYMENT_LABELS,
  type PaymentStatus,
} from "@/lib/orders";
import {
  saveRetailOrder,
  getRetailOrder,
  listRetailOrders,
  createRetailOrder,
  type RetailItem,
  type SavedRetailOrder,
} from "@/lib/retail-orders";
import { sendRetailOrderEmail } from "@/lib/retail-notify";
import { generateRetailPdf } from "@/lib/retail-pdf";
import {
  RETAIL_STATUSES,
  MAT_TYPES,
  INNER_MAT_TYPES,
  GLASS_TYPES,
  PRINT_TYPES,
  toMM,
  type RetailStatus,
} from "@/data/perakende";
// Websitedeki hesaplayıcıyla AYNI fiyat çekirdeği: sunucu, istemcinin
// gönderdiği kalem tutarlarını yeniden hesaplayıp doğrular ve üretim
// kurallarını (tabaka/cam/dış ölçü sınırları) uygular.
import {
  hesaplaPerakende,
  dogrulaPerakende,
  type PerakendeGirdi,
} from "@/lib/perakende-fiyat";
import { kurus } from "@/lib/num";

export const dynamic = "force-dynamic";

const r2 = (n: number) => Math.round((Number(n) || 0) * 100) / 100;
const s = (v: unknown, max = 200) => String(v ?? "").slice(0, max);

function sanitizeItems(raw: unknown): RetailItem[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((it: any): RetailItem => ({
      artWidth: Number(it?.artWidth) || 0,
      artWidthUnit: it?.artWidthUnit === "mm" ? "mm" : "cm",
      artHeight: Number(it?.artHeight) || 0,
      artHeightUnit: it?.artHeightUnit === "mm" ? "mm" : "cm",
      frameCode: s(it?.frameCode, 60),
      framePriceTL: r2(it?.framePriceTL),
      manualPrice: Boolean(it?.manualPrice),
      matType: s(it?.matType, 60),
      matCode: s(it?.matCode, 20),
      matColor: s(it?.matColor, 20),
      matColorHex: s(it?.matColorHex, 20),
      doubleMat: Boolean(it?.doubleMat),
      innerMatType: s(it?.innerMatType, 60),
      innerMatColor: s(it?.innerMatColor, 20),
      innerMatColorHex: s(it?.innerMatColorHex, 20),
      altMontaj: s(it?.altMontaj, 10),
      zeminEnabled: Boolean(it?.zeminEnabled),
      zeminType: s(it?.zeminType, 60),
      zeminColor: s(it?.zeminColor, 20),
      zeminColorHex: s(it?.zeminColorHex, 20),
      matTop: Number(it?.matTop) || 0,
      matRight: Number(it?.matRight) || 0,
      matBottom: Number(it?.matBottom) || 0,
      matLeft: Number(it?.matLeft) || 0,
      pencereSayisi: Math.min(9, Math.max(1, Math.round(Number(it?.pencereSayisi) || 1))),
      pencereDuzen: s(it?.pencereDuzen, 8),
      pencereAralik: r2(it?.pencereAralik),
      kasa: Boolean(it?.kasa),
      glassType: s(it?.glassType, 60),
      printType: s(it?.printType, 60),
      frameCost: r2(it?.frameCost),
      matCost: r2(it?.matCost),
      glassCost: r2(it?.glassCost),
      printCost: r2(it?.printCost),
      itemTotal: r2(it?.itemTotal),
    }))
    .filter((it) => it.artWidth > 0 && it.artHeight > 0);
}

export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const date = req.nextUrl.searchParams.get("date");
  const range = req.nextUrl.searchParams.get("range") || "today";
  let keys: string[];
  if (date && /^\d{4}-\d{2}-\d{2}$/.test(date)) keys = [date];
  else if (range === "week") keys = lastNDateKeys(7);
  else keys = [istanbulDateKey()];

  const orders = await listRetailOrders(keys);
  return NextResponse.json({ orders, blob: blobConfigured() });
}

export async function POST(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const body = await req.json().catch(() => null);
  if (!body) return NextResponse.json({ error: "Geçersiz istek" }, { status: 400 });

  const items = sanitizeItems(body.items);
  if (items.length === 0) {
    return NextResponse.json({ error: "Sipariş kalemi yok" }, { status: 400 });
  }
  const customerName = s(body.customerName, 120).trim();
  const customerPhone = s(body.customerPhone, 40).trim();
  if (!customerName || !customerPhone) {
    return NextResponse.json(
      { error: "Müşteri adı ve telefonu zorunludur" },
      { status: 400 }
    );
  }

  // ---- Sunucu tarafı fiyat doğrulama (websitedeki hesaplayıcıyla aynı) ----
  // Her kalem çekirdekle yeniden hesaplanır; üretim kuralları (paspartu
  // tabakası 80×120, cam plakası 100×140, dış ölçü 290 cm...) burada da
  // uygulanır. İstemcinin tutarı ±1 TL içinde tutmalı — tutmuyorsa istemci
  // eski sürümdür, sayfa yenilenmelidir.
  const usdRateNum = r2(body.usdRate);
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const adres = `${i + 1}. kalem`;
    if (!(it.framePriceTL > 0)) {
      return NextResponse.json(
        { error: `${adres}: çerçeve metre fiyatı eksik.` },
        { status: 400 }
      );
    }
    const matSec = MAT_TYPES.find((m) => m.name === it.matType);
    const glassSec = GLASS_TYPES.find((g) => g.name === it.glassType);
    const printSec = PRINT_TYPES.find((p) => p.name === it.printType);
    if (!matSec || !glassSec || !printSec) {
      return NextResponse.json(
        { error: `${adres}: geçersiz paspartu/cam/baskı türü.` },
        { status: 400 }
      );
    }
    const innerSec = it.doubleMat
      ? INNER_MAT_TYPES.find((m) => m.name === it.innerMatType)
      : undefined;
    if (it.doubleMat && !innerSec) {
      return NextResponse.json(
        { error: `${adres}: geçersiz iç paspartu türü.` },
        { status: 400 }
      );
    }
    const zeminSec = it.zeminEnabled
      ? INNER_MAT_TYPES.find((m) => m.name === it.zeminType)
      : undefined;
    if (it.zeminEnabled && !zeminSec) {
      return NextResponse.json(
        { error: `${adres}: geçersiz zemin türü.` },
        { status: 400 }
      );
    }
    if (printSec.usdPerM2 > 0 && !(usdRateNum > 0)) {
      return NextResponse.json(
        { error: `${adres}: baskı fiyatı için USD kuru gerekli.` },
        { status: 400 }
      );
    }
    const duzenParca = /^(\d)x(\d)$/.exec(it.pencereDuzen || "");
    const girdi: PerakendeGirdi = {
      wMM: toMM(it.artWidth, it.artWidthUnit),
      hMM: toMM(it.artHeight, it.artHeightUnit),
      kenar: { ust: it.matTop, alt: it.matBottom, sol: it.matLeft, sag: it.matRight },
      matPrice: matSec.price,
      doubleMat: it.doubleMat,
      innerMatPrice: innerSec?.price || 0,
      icSeritMm: Number(String(it.altMontaj).replace(",", ".")) || 5,
      zeminEnabled: it.zeminEnabled,
      zeminPrice: zeminSec?.price || 0,
      pencereSayisi: it.pencereSayisi || 1,
      pencereRows: duzenParca ? Number(duzenParca[1]) : undefined,
      pencereCols: duzenParca ? Number(duzenParca[2]) : undefined,
      pencereAralikMm: it.pencereAralik || undefined,
      camPrice: glassSec.price,
      camMaxKisaMM: glassSec.maxKisaMM,
      camMaxUzunMM: glassSec.maxUzunMM,
      camUyariKisaMM: glassSec.uyariKisaMM,
      camUyariUzunMM: glassSec.uyariUzunMM,
      kasa: !!it.kasa,
      framePriceTL: it.framePriceTL,
      printUsdPerM2: printSec.usdPerM2,
      usdRate: usdRateNum,
    };
    const engel = dogrulaPerakende(girdi).filter((h) => h.seviye === "engel");
    if (engel.length) {
      return NextResponse.json(
        { error: `${adres}: ${engel[0].mesaj}` },
        { status: 400 }
      );
    }
    const hesap = hesaplaPerakende(girdi);
    const sunucuToplam = kurus(
      kurus(hesap.frameCost) +
        kurus(hesap.matCost) +
        kurus(hesap.glassCost) +
        kurus(hesap.printCost)
    );
    if (Math.abs(sunucuToplam - it.itemTotal) > 1) {
      return NextResponse.json(
        {
          error: `${adres}: fiyat sunucu hesabıyla uyuşmuyor (₺${sunucuToplam} hesaplandı, ₺${it.itemTotal} gönderildi). Sayfayı yenileyip tekrar deneyin.`,
        },
        { status: 400 }
      );
    }
    // Sunucu hesabı esas alınır — kayda giden döküm her zaman çekirdeğin sonucu
    it.frameCost = kurus(hesap.frameCost);
    it.matCost = kurus(hesap.matCost);
    it.glassCost = kurus(hesap.glassCost);
    it.printCost = kurus(hesap.printCost);
    it.itemTotal = sunucuToplam;
  }

  const gross = r2(items.reduce((sum, it) => sum + it.itemTotal, 0));
  const discount = Math.min(r2(body.discount), gross);
  const total = r2(gross - Math.max(0, discount));

  // Kapora — 0 ile toplam arasına kıstırılır; tahsilat kaydına da işlenir
  const kaporaTutar = Math.min(total, Math.max(0, r2(body?.kapora?.amount)));
  const kaporaYontem: "nakit" | "krediKarti" =
    body?.kapora?.method === "krediKarti" ? "krediKarti" : "nakit";

  const now = new Date();
  // ---- Önce KALICI KAYIT, sonra bildirim ----
  // Numara çakışmasız üretilir (retail/no rezervasyonu); kayıt başarısızsa
  // e-posta/PDF hiç üretilmez — "bildirimde var, panelde yok" sipariş kalmaz.
  const taslak: Omit<SavedRetailOrder, "orderId"> = {
    dateKey: istanbulDateKey(now),
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
    status: "Beklemede",
    employee: user.name,
    customerName,
    customerPhone,
    customerEmail: s(body.customerEmail, 120).trim(),
    customerAddress: s(body.customerAddress, 240).trim(),
    customerId: s(body.customerId, 40),
    branch: body.branch === "istanbul" ? "istanbul" : "ankara",
    payment:
      kaporaTutar <= 0
        ? "bekliyor"
        : kaporaTutar + 0.01 >= total
          ? "odendi"
          : "kismi",
    paidAmount: kaporaTutar,
    usdRate: r2(body.usdRate),
    deliveryDate: s(body.deliveryDate, 20),
    notes: s(body.notes, 1000),
    items,
    gross,
    discount: Math.max(0, discount),
    total,
  };

  let orderId = "";
  let stored = false;
  try {
    ({ orderId, stored } = await createRetailOrder(taslak));
  } catch (err) {
    console.error("Perakende siparişi kaydedilemedi:", err);
  }
  if (blobConfigured() && !stored) {
    return NextResponse.json(
      {
        error:
          "Sipariş KAYDEDİLEMEDİ (depo hatası). E-posta gönderilmedi — bilgiler formda duruyor, lütfen tekrar deneyin.",
      },
      { status: 503 }
    );
  }
  const order: SavedRetailOrder = { ...taslak, orderId };

  // Kapora tahsilat kaydı — kasa ve cari hareketler eksiksiz kalsın.
  // Tahsilat yazılamazsa sipariş kaydı geçerli kalır (listeden düzeltilebilir).
  if (stored && kaporaTutar > 0) {
    try {
      const { saveTahsilat, newTahsilatId, istanbulDateKey: gunKey } = await import(
        "@/lib/tahsilat"
      );
      const { applyTahsilatDelta } = await import("@/lib/finans-ozet");
      const t = {
        id: newTahsilatId(),
        dateKey: gunKey(),
        createdAt: new Date().toISOString(),
        createdBy: user.name,
        branch: (order.branch || "ankara") as "ankara" | "istanbul",
        customerId: order.customerId || undefined,
        customerName: order.customerName,
        orderId: order.orderId,
        orderDateKey: order.dateKey,
        amount: kaporaTutar,
        currency: "TL" as const,
        method: kaporaYontem,
        note: "Perakende kapora",
        kaynak: "panel" as const,
      };
      await saveTahsilat(t);
      await applyTahsilatDelta(t, 1);
    } catch (err) {
      console.error("Kapora tahsilatı kaydedilemedi:", err);
    }
  }

  // Üretim PDF'i — e-posta ekinde gider, ayrıca listeden indirilebilir
  let pdf: Buffer | undefined;
  try {
    pdf = await generateRetailPdf(order);
  } catch {
    /* PDF hatası siparişi engellemesin */
  }

  let emailSent = false;
  try {
    emailSent = await sendRetailOrderEmail(order, pdf);
  } catch {
    /* e-posta hatası siparişi engellemesin */
  }

  return NextResponse.json({
    ok: true,
    orderId,
    dateKey: order.dateKey,
    saved: stored,
    emailSent,
    kapora: kaporaTutar,
    kalan: r2(total - kaporaTutar),
  });
}

export async function PATCH(req: NextRequest) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ error: "Yetkisiz" }, { status: 401 });
  }

  const d = req.nextUrl.searchParams.get("d") || "";
  const id = req.nextUrl.searchParams.get("id") || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^PRK-[A-Z0-9-]+$/i.test(id)) {
    return NextResponse.json({ error: "Geçersiz sipariş" }, { status: 400 });
  }

  const body = await req.json().catch(() => null);
  const order = await getRetailOrder(d, id);
  if (!order) return NextResponse.json({ error: "Sipariş bulunamadı" }, { status: 404 });
  const oncekiPaidAmount = Number(order.paidAmount) || 0;

  if (body?.status !== undefined) {
    const status = body.status as RetailStatus;
    if (!RETAIL_STATUSES.includes(status)) {
      return NextResponse.json({ error: "Geçersiz durum" }, { status: 400 });
    }
    order.status = status;
  }

  // Ödeme (cari) güncelleme
  if (body?.payment !== undefined) {
    const payment = String(body.payment) as PaymentStatus;
    if (!(payment in PAYMENT_LABELS)) {
      return NextResponse.json({ error: "Geçersiz ödeme durumu" }, { status: 400 });
    }
    order.payment = payment;
    if (payment === "odendi") order.paidAmount = order.total;
    else if (payment === "bekliyor") order.paidAmount = 0;
  }
  if (body?.paidAmount !== undefined) {
    const paid = Math.max(0, r2(body.paidAmount));
    order.paidAmount = paid;
    if (paid <= 0) order.payment = "bekliyor";
    else if (paid + 0.01 >= order.total) {
      order.payment = "odendi";
      order.paidAmount = order.total;
    } else order.payment = "kismi";
  }

  // Ödeme artışı tahsilat kaydı üretir — kasa ve cari hareketler eksiksiz
  // kalsın (azalış = düzeltme, kayıt üretmez).
  const delta = r2((Number(order.paidAmount) || 0) - oncekiPaidAmount);
  if (delta > 0) {
    try {
      const { saveTahsilat, newTahsilatId, istanbulDateKey } = await import(
        "@/lib/tahsilat"
      );
      const { applyTahsilatDelta } = await import("@/lib/finans-ozet");
      const t = {
        id: newTahsilatId(),
        dateKey: istanbulDateKey(),
        createdAt: new Date().toISOString(),
        createdBy: user.name,
        branch: (order.branch || "ankara") as "ankara" | "istanbul",
        customerId: order.customerId || undefined,
        customerName: order.customerName,
        orderId: order.orderId,
        orderDateKey: order.dateKey,
        amount: delta,
        currency: "TL" as const,
        method: "diger" as const,
        note: "Perakende listesinden ödeme durumu değişikliği",
        kaynak: "panel" as const,
      };
      await saveTahsilat(t);
      await applyTahsilatDelta(t, 1);
    } catch {
      /* tahsilat kaydı üretilemese de sipariş güncellemesi geçerli kalır */
    }
  }

  order.updatedAt = new Date().toISOString();
  await saveRetailOrder(order);
  return NextResponse.json({
    ok: true,
    status: order.status,
    payment: order.payment,
    paidAmount: order.paidAmount,
  });
}
