import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import {
  sendOrderEmail,
  sendOrderWhatsApp,
  waLink,
  type OrderPayload,
  type OrderLine,
} from "@/lib/notify";
import {
  createOrder,
  getOrder,
  listOrders,
  listAllOrders,
  lastNDateKeys,
  istanbulDateKey,
  sanitizeLines,
  computeTotals,
  getDailyRates,
  saveDailyRates,
  blobConfigured,
  readOrderIndex,
  rebuildOrderIndexFrom,
  type SavedOrder,
  type OrderIndexEntry,
} from "@/lib/orders";
import { isKurYetkili } from "@/data/users";
import { eslesir } from "@/lib/search-norm";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";


// Sipariş listesi:
//   ?range=today | yesterday | week   ?date=YYYY-MM-DD
//   ?q=<metin>            → tüm siparişlerde arama (müşteri / sipariş no / çalışan / not)
//   ?musteri=<customerId> → o müşterinin son siparişleri (mükerrer sipariş uyarısı)
export async function GET(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const date = url.searchParams.get("date");
  const range = url.searchParams.get("range") || "today";
  const q = (url.searchParams.get("q") || "").trim();
  const musteri = (url.searchParams.get("musteri") || "").trim();
  const musteriAd = (url.searchParams.get("musteriAd") || "").trim();

  // Müşteri bazlı son siparişler — sipariş formundaki mükerrer uyarısı için.
  // Defterden seçildiyse customerId, seçilmediyse yazılan ada göre eşleşir.
  // Aylık indeks üzerinden çalışır: tüm siparişleri tek tek okumak yerine
  // yalnızca eşleşenler getirilir. İndeks boşsa tam taramaya düşülür ve
  // indeks o taramadan kurulur (kendi kendini onarır).
  if (musteri || musteriAd) {
    const gun = Math.min(30, Math.max(1, Number(url.searchParams.get("gun")) || 7));
    const izin = new Set(lastNDateKeys(gun));
    const uygun = (o: { dateKey: string; status: string; customerId?: string; customer: string }) =>
      izin.has(o.dateKey) &&
      // İptal edilen sipariş mükerrer uyarısına girmez
      o.status !== "iptal" &&
      (musteri ? o.customerId === musteri : eslesir(musteriAd, o.customer));

    const idx = await readOrderIndex(2); // son 30 gün en fazla 2 aya yayılır
    let orders: SavedOrder[];
    if (idx.length) {
      const hits = idx.filter(uygun).slice(0, 50);
      orders = (
        await Promise.all(hits.map((e) => getOrder(e.dateKey, e.orderId)))
      ).filter(Boolean) as SavedOrder[];
      orders.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
    } else {
      const hepsi = await listAllOrders();
      await rebuildOrderIndexFrom(hepsi);
      orders = hepsi.filter(uygun);
    }
    return NextResponse.json({ ok: true, orders });
  }

  // Serbest arama — müşteri adı, sipariş numarası, çalışan veya not içinde.
  // Önce indeksten eşleşen (gün, no) bulunur, sonra yalnız onlar okunur.
  if (q) {
    const uygun = (o: OrderIndexEntry | SavedOrder) =>
      eslesir(q, o.customer, o.orderId, o.employee, o.note);
    const idx = await readOrderIndex();
    let orders: SavedOrder[];
    if (idx.length) {
      const hits = idx.filter(uygun).slice(0, 200);
      orders = (
        await Promise.all(hits.map((e) => getOrder(e.dateKey, e.orderId)))
      ).filter(Boolean) as SavedOrder[];
      orders.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
    } else {
      const hepsi = await listAllOrders();
      await rebuildOrderIndexFrom(hepsi);
      orders = hepsi.filter(uygun).slice(0, 200);
    }
    return NextResponse.json({ ok: true, orders, arama: q });
  }

  let dateKeys: string[];
  if (date && /^\d{4}-\d{2}-\d{2}$/.test(date)) dateKeys = [date];
  else if (range === "week") dateKeys = lastNDateKeys(7);
  else if (range === "days15") dateKeys = lastNDateKeys(15);
  else if (range === "yesterday") dateKeys = [lastNDateKeys(2)[1]];
  else dateKeys = lastNDateKeys(1);

  const orders = await listOrders(dateKeys);
  return NextResponse.json({ ok: true, orders });
}

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }

  const body = await req.json().catch(() => null);
  const lines = sanitizeLines(body?.lines);
  if (!lines.length) {
    return NextResponse.json({ ok: false, error: "Geçerli satır yok." }, { status: 400 });
  }

  // ---- Sunucu tarafı doğrulama (poka-yoke) ----
  // İstemci hesabına körü körüne güvenme: ₺0 birim fiyatlı satır (kur boşken
  // USD×0 gibi) veya negatif tutar kalıcı kayda hiç girmesin.
  for (const l of lines) {
    if (!(l.unitPriceTL > 0) || !(l.lineTotal > 0)) {
      return NextResponse.json(
        {
          ok: false,
          error: `"${l.name}" satırının fiyatı geçersiz (₺${l.unitPriceTL}). Kur girilmemiş olabilir — kuru kontrol edip tekrar deneyin.`,
        },
        { status: 400 }
      );
    }
  }
  const discountPct = Math.max(0, Number(body.discountPct) || 0);
  if (discountPct > 100) {
    return NextResponse.json({ ok: false, error: "İskonto %100'ü aşamaz." }, { status: 400 });
  }

  const now = new Date();
  const dateKey = istanbulDateKey(now);
  const rate = Number(body.rate) || 0;
  const euroRate = Number(body.euroRate) || 0;

  // Kur kilidi sunucuda da geçerli: günün kuru yetkili tarafından
  // sabitlendiyse, yetkisiz kullanıcıdan farklı kurla gelen istek reddedilir
  // (ekrandaki kilit aşılsa ya da form eski kurla açık kalsa bile).
  try {
    const gunKuru = await getDailyRates(dateKey);
    if (gunKuru?.sabit && !isKurYetkili(user.username)) {
      const farkli =
        (gunKuru.rate > 0 && rate > 0 && Math.abs(rate - gunKuru.rate) > 0.005) ||
        (gunKuru.euroRate > 0 && euroRate > 0 && Math.abs(euroRate - gunKuru.euroRate) > 0.005);
      if (farkli) {
        return NextResponse.json(
          {
            ok: false,
            error: `Günün kuru yetkili tarafından ₺${gunKuru.rate} olarak belirlendi. Sayfayı yenileyip güncel kurla tekrar deneyin.`,
          },
          { status: 409 }
        );
      }
    }
  } catch {
    /* kur okunamazsa siparişi engelleme */
  }

  const vatApplied = !!body.vatApplied;
  const { gross, discount, vatAmount, net } = computeTotals(lines, discountPct, vatApplied);

  // ---- Önce KALICI KAYIT, sonra bildirim ----
  // Numara çakışmasız üretilir (orders/no rezervasyonu); kayıt başarısızsa
  // hiçbir bildirim gitmez — "bildirimde var, panelde yok" sipariş kalmaz.
  const taslak: Omit<SavedOrder, "orderId"> = {
    dateKey,
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
    status: "olusturuldu",
    employee: user.name,
    customer: String(body.customer || "").slice(0, 200),
    // Müşteri defterinden seçildiyse cari takip için bağlanır
    customerId: String(body.customerId || "").slice(0, 40),
    branch: body.branch === "istanbul" ? "istanbul" : "ankara",
    payment: "bekliyor",
    paidAmount: 0,
    note: String(body.note || "").slice(0, 500),
    rate,
    euroRate,
    discountPct,
    vatApplied,
    lines,
    gross,
    discount,
    vatAmount,
    net,
    rows: Array.isArray(body.rows) ? body.rows.slice(0, 100) : undefined,
  };

  let orderId = "";
  let stored = false;
  try {
    ({ orderId, stored } = await createOrder(taslak));
  } catch (err) {
    console.error("Sipariş kaydedilemedi:", err);
  }
  if (blobConfigured() && !stored) {
    return NextResponse.json(
      {
        ok: false,
        error:
          "Sipariş KAYDEDİLEMEDİ (depo hatası). Bildirim gönderilmedi — bilgiler formda duruyor, lütfen tekrar deneyin.",
      },
      { status: 503 }
    );
  }

  const order: OrderPayload = {
    orderId,
    employee: taslak.employee,
    customer: taslak.customer,
    note: taslak.note,
    rate,
    euroRate,
    discountPct,
    vatApplied,
    lines,
    gross,
    discount,
    vatAmount,
    net,
    dateStr: now.toLocaleString("tr-TR", { timeZone: "Europe/Istanbul" }),
  };
  const saved: SavedOrder = { ...taslak, orderId };

  // Günün kuru daha önce kaydedilmediyse bu siparişteki kuru günlük kur yap
  if (order.rate > 0) {
    try {
      const existing = await getDailyRates(saved.dateKey);
      if (!existing) {
        await saveDailyRates(saved.dateKey, {
          rate: order.rate,
          euroRate: order.euroRate,
          updatedAt: now.toISOString(),
          by: user.name,
        });
      }
    } catch (err) {
      console.error("Günlük kur kaydedilemedi:", err);
    }
  }

  let emailSent = false;
  let waSent = false;
  let patronWa = false;
  let pdf: Buffer | undefined;
  try {
    const { generateOrderPdf } = await import("@/lib/order-pdf");
    pdf = await generateOrderPdf(order);
  } catch (err) {
    console.error("PDF üretilemedi:", err);
  }
  try {
    emailSent = await sendOrderEmail(order, pdf);
  } catch (err) {
    console.error("E-posta gönderilemedi:", err);
  }
  try {
    waSent = await sendOrderWhatsApp(order);
  } catch (err) {
    console.error("WhatsApp gönderilemedi:", err);
  }
  // Patrona fiş PDF'i (WhatsApp belge mesajı) — hata siparişi asla engellemez
  if (pdf) {
    try {
      const { patronBildirimHazir, sendPdfToPatron } = await import("@/lib/whatsapp-pdf");
      if (patronBildirimHazir()) {
        const r = await sendPdfToPatron(pdf, {
          tur: "toptan",
          orderId: order.orderId,
          musteri: order.customer,
          tutar: order.net,
          calisan: order.employee,
        });
        patronWa = r.ok;
        if (r.hatalar.length) console.error("Patron WhatsApp:", r.hatalar.join(" | "));
        if (r.notlar.length) console.warn("Patron WhatsApp:", r.notlar.join(" | "));
      }
    } catch (err) {
      console.error("Patron WhatsApp gönderilemedi:", err);
    }
  }

  // Müşteriye sipariş bildirimi — form işaretliyse ve müşteri defterden
  // seçilmişse. Önce WhatsApp (fiş PDF'i, onaylı şablon); WhatsApp kurulu
  // değilse ya da Meta anında reddederse SMS. Meta kabul edip sonradan
  // "teslim edilemedi" derse (numarada WhatsApp yok) webhook SMS'e düşer.
  // Hiçbir bildirim hatası sipariş kaydını geri döndürmez.
  let smsSent = false;
  let smsInfo = "";
  let musteriWa = false;
  if (body.sendSms && saved.customerId) {
    try {
      const { getCustomer } = await import("@/lib/customers");
      const c = await getCustomer(saved.customerId);
      if (!c?.phone) {
        smsInfo = "Müşterinin kayıtlı telefonu yok.";
      } else {
        const { stripTurkish } = await import("@/lib/sms-format");
        const mesaj = stripTurkish(
          `Sayin musterimiz, ${order.orderId} numarali siparisiniz alinmistir. Tesekkur ederiz. Olga Cerceve`
        );
        let smsGerekli = true;

        if (pdf) {
          try {
            const { musteriBildirimHazir, sendPdfToCustomer } = await import("@/lib/whatsapp-pdf");
            if (musteriBildirimHazir()) {
              const w = await sendPdfToCustomer(pdf, { telefon: c.phone, musteri: order.customer, orderId: order.orderId });
              if (w.ok && w.wamid && w.to) {
                musteriWa = true;
                smsGerekli = false;
                const { bekleyenKaydet } = await import("@/lib/wa-bekleyen");
                await bekleyenKaydet({
                  wamid: w.wamid, orderId: order.orderId, dateKey: saved.dateKey, tur: "toptan",
                  telefon: w.to, musteri: order.customer, smsMetni: mesaj, createdAt: now.toISOString(),
                }).catch((err) => console.error("WhatsApp bekleyen kaydı yazılamadı:", err));
              } else {
                console.warn("Müşteri WhatsApp reddedildi, SMS'e düşülüyor:", w.hata);
              }
            }
          } catch (err) {
            console.error("Müşteri WhatsApp gönderilemedi, SMS'e düşülüyor:", err);
          }
        }

        if (smsGerekli) {
          const { smsConfigured, sendSms } = await import("@/lib/sms");
          if (smsConfigured()) {
            const r = await sendSms([c.phone], mesaj);
            smsSent = r.ok;
            if (!r.ok) smsInfo = r.error || "SMS gönderilemedi.";

            const { saveSmsRecord, newSmsId, istanbulDateKey: gun } = await import("@/lib/sms-log");
            const t = new Date();
            await saveSmsRecord({
              id: newSmsId(t),
              dateKey: gun(t),
              createdAt: t.toISOString(),
              sender: user.name,
              message: mesaj,
              recipients: r.sent,
              segments: 1,
              credits: r.sent.length,
              ok: r.ok,
              jobId: r.jobId,
              error: r.error,
              iysfilter: "0",
            }).catch(() => {});
          }
        }
      }
    } catch (err) {
      console.error("Müşteri bildirimi gönderilemedi:", err);
      smsInfo = "Bildirim gönderiminde hata oluştu.";
    }
  }

  return NextResponse.json({
    ok: true,
    orderId: order.orderId,
    dateKey: saved.dateKey,
    stored,
    emailSent,
    waSent,
    patronWa,
    waLink: waSent ? undefined : waLink(order),
    smsSent,
    smsInfo,
    musteriWa,
    net,
  });
}
