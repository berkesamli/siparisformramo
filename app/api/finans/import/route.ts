import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import {
  saveTahsilat,
  type Tahsilat,
  type TahsilatYontem,
  type ParaBirimi,
} from "@/lib/tahsilat";
import { saveAcilisBakiye, type AcilisBakiye } from "@/lib/acilis-bakiye";
import { saveCekSenet, type CekSenet, type CekSenetDurum } from "@/lib/ceksenet";
import { saveGider, type Gider, type GiderYontem } from "@/lib/gider";
import { rebuildOzet } from "@/lib/finans-ozet";
import { kurus } from "@/lib/num";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Excel aktarımı ve düzeltmeler için toplu veri kapısı — yalnızca sahipler.
//
// Kayıtlar deterministik id taşır: aynı batch tekrar gönderilirse üzerine
// yazar, mükerrer oluşturmaz (idempotent). Desteklenen türler:
//   acilis    → finans/acilis/<customerId>.json
//   tahsilat  → finans/tahsilat/<ay>/<id>.json  (id verilmek zorunda)
//   rebuild   → verilen ayların özetini kaynak kayıtlardan yeniden kurar
// Faz 2'de: ceksenet, gider, order-patch eklenecek.

const MAX_BATCH = 100;
const YONTEMLER: TahsilatYontem[] = [
  "nakit",
  "havale",
  "krediKarti",
  "cek",
  "senet",
  "diger",
];
const BIRIMLER: ParaBirimi[] = ["TL", "USD", "EUR"];

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as {
    kind?: string;
    records?: unknown[];
    months?: string[];
  } | null;
  if (!body?.kind) {
    return NextResponse.json({ ok: false, error: "kind gerekli." }, { status: 400 });
  }

  if (body.kind === "rebuild") {
    const months = (body.months || []).filter((m) => /^\d{4}-\d{2}$/.test(m));
    for (const m of months) await rebuildOzet(m);
    return NextResponse.json({ ok: true, rebuilt: months });
  }

  const records = Array.isArray(body.records) ? body.records : [];
  if (!records.length || records.length > MAX_BATCH) {
    return NextResponse.json(
      { ok: false, error: `records 1-${MAX_BATCH} arası olmalı.` },
      { status: 400 }
    );
  }

  let yazilan = 0;
  const hatalar: string[] = [];

  if (body.kind === "acilis") {
    for (const raw of records as Partial<AcilisBakiye>[]) {
      const customerId = String(raw.customerId || "").slice(0, 40);
      const customerName = String(raw.customerName || "").trim().slice(0, 200);
      const asOf = String(raw.asOf || "");
      if (!customerId || !customerName || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) {
        hatalar.push(`acilis: eksik alan (${customerName || customerId || "?"})`);
        continue;
      }
      const a: AcilisBakiye = {
        customerId,
        customerName,
        branch: raw.branch === "istanbul" ? "istanbul" : "ankara",
        amount: kurus(Number(raw.amount) || 0),
        asOf,
        note: String(raw.note || "").slice(0, 300) || undefined,
        kaynak: "excel",
        createdAt: new Date().toISOString(),
        createdBy: user.name,
      };
      if (await saveAcilisBakiye(a)) yazilan++;
    }
  } else if (body.kind === "tahsilat") {
    for (const raw of records as Partial<Tahsilat>[]) {
      const id = String(raw.id || "");
      const dateKey = String(raw.dateKey || "");
      const customerName = String(raw.customerName || "").trim().slice(0, 200);
      const amount = kurus(Number(raw.amount) || 0);
      const method = (raw.method || "diger") as TahsilatYontem;
      const currency = (raw.currency || "TL") as ParaBirimi;
      if (
        !/^T-[\dA-Za-z-]+$/.test(id) ||
        !/^\d{4}-\d{2}-\d{2}$/.test(dateKey) ||
        !customerName ||
        amount <= 0 ||
        !YONTEMLER.includes(method) ||
        !BIRIMLER.includes(currency)
      ) {
        hatalar.push(`tahsilat: geçersiz kayıt (${id || customerName || "?"})`);
        continue;
      }
      const t: Tahsilat = {
        id,
        dateKey,
        createdAt: new Date().toISOString(),
        createdBy: user.name,
        branch: raw.branch === "istanbul" ? "istanbul" : "ankara",
        customerId: String(raw.customerId || "").slice(0, 40) || undefined,
        customerName,
        orderId: undefined, // geçmiş Excel tahsilatları siparişe bağlanmaz
        orderDateKey: undefined,
        amount,
        currency,
        method,
        tahsilEden:
          String(raw.tahsilEden || "").trim().slice(0, 100) || undefined,
        note: String(raw.note || "").slice(0, 300) || undefined,
        kaynak: raw.kaynak === "migrasyon" ? "migrasyon" : "excel",
      };
      if (await saveTahsilat(t)) yazilan++;
      // Özet deltası burada uygulanmaz — toplu aktarım sonrası "rebuild"
      // çağrısı tüm ayları tek seferde doğru kurar (daha az yazma, daha
      // güvenli sonuç).
    }
  } else if (body.kind === "ceksenet") {
    const DURUMLAR: CekSenetDurum[] = [
      "portfoyde",
      "tahsil",
      "ciro",
      "karsiliksiz",
      "odendi",
      "iade",
    ];
    for (const raw of records as Partial<CekSenet>[]) {
      const id = String(raw.id || "");
      const tutar = kurus(Number(raw.tutar) || 0);
      const vade = String(raw.vade || "");
      const durum = (raw.durum || "portfoyde") as CekSenetDurum;
      if (
        !/^CS-[\dA-Za-z-]+$/.test(id) ||
        tutar <= 0 ||
        !/^\d{4}-\d{2}-\d{2}$/.test(vade) ||
        !DURUMLAR.includes(durum)
      ) {
        hatalar.push(`ceksenet: geçersiz kayıt (${id || raw.customerName || "?"})`);
        continue;
      }
      const now = new Date().toISOString();
      const cs: CekSenet = {
        id,
        createdAt: now,
        createdBy: user.name,
        tur: raw.tur === "verilen" ? "verilen" : "alinan",
        kind: raw.kind === "senet" ? "senet" : "cek",
        branch: raw.branch === "istanbul" ? "istanbul" : "ankara",
        banka: String(raw.banka || "").slice(0, 100) || undefined,
        bankaSube: String(raw.bankaSube || "").slice(0, 100) || undefined,
        hesapNo: String(raw.hesapNo || "").slice(0, 60) || undefined,
        belgeNo: String(raw.belgeNo || "").slice(0, 60) || undefined,
        cekSahibi: String(raw.cekSahibi || "").slice(0, 200) || undefined,
        tutar,
        vade,
        customerId: String(raw.customerId || "").slice(0, 40) || undefined,
        customerName: String(raw.customerName || "").slice(0, 200) || undefined,
        supplier: String(raw.supplier || "").slice(0, 200) || undefined,
        durum,
        tahsilDate: String(raw.tahsilDate || "").slice(0, 10) || undefined,
        ciroTarget: String(raw.ciroTarget || "").slice(0, 200) || undefined,
        history: Array.isArray(raw.history) && raw.history.length
          ? raw.history
          : [
              {
                durum,
                date: String(raw.tahsilDate || vade).slice(0, 10),
                by: user.name,
                note: "Excel aktarımı",
              },
            ],
        note: String(raw.note || "").slice(0, 300) || undefined,
        kaynak: "excel",
      };
      // Excel aktarımında otomatik Tahsilat ÜRETİLMEZ — geçmiş cari zaten
      // devir bakiyesinde kapanmıştır; çift sayım olmasın.
      if (await saveCekSenet(cs)) yazilan++;
    }
  } else if (body.kind === "gider") {
    const YONTEMLER: GiderYontem[] = ["nakit", "havale", "krediKarti", "cek", "diger"];
    for (const raw of records as Partial<Gider>[]) {
      const id = String(raw.id || "");
      const dateKey = String(raw.dateKey || "");
      const amount = kurus(Number(raw.amount) || 0);
      const method = (raw.method || "diger") as GiderYontem;
      const category = String(raw.category || "").trim().slice(0, 60);
      if (
        !/^G-[\dA-Za-z-]+$/.test(id) ||
        !/^\d{4}-\d{2}-\d{2}$/.test(dateKey) ||
        amount <= 0 ||
        !category ||
        !YONTEMLER.includes(method)
      ) {
        hatalar.push(`gider: geçersiz kayıt (${id || category || "?"})`);
        continue;
      }
      const g: Gider = {
        id,
        dateKey,
        createdAt: new Date().toISOString(),
        createdBy: user.name,
        branch: raw.branch === "istanbul" ? "istanbul" : "ankara",
        category: category.toLocaleLowerCase("tr-TR"),
        description: String(raw.description || "").slice(0, 300),
        amount,
        currency: raw.currency === "USD" || raw.currency === "EUR" ? raw.currency : "TL",
        method,
        supplier: String(raw.supplier || "").slice(0, 200) || undefined,
        personelId: String(raw.personelId || "").slice(0, 40) || undefined,
        note: String(raw.note || "").slice(0, 300) || undefined,
        kaynak: "excel",
      };
      if (await saveGider(g)) yazilan++;
      // Özet, aktarım sonrası "rebuild" ile kurulur.
    }
  } else {
    return NextResponse.json(
      { ok: false, error: `Bilinmeyen kind: ${body.kind}` },
      { status: 400 }
    );
  }

  return NextResponse.json({ ok: true, yazilan, hatalar });
}
