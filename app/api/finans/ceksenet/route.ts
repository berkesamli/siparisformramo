import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import {
  saveCekSenet,
  getCekSenet,
  listCekSenet,
  newCekSenetId,
  allowedTransitions,
  type CekSenet,
  type CekSenetDurum,
} from "@/lib/ceksenet";
import {
  saveTahsilat,
  deleteTahsilat,
  getTahsilat,
  newTahsilatId,
  istanbulDateKey,
  ayKey,
  type Tahsilat,
} from "@/lib/tahsilat";
import {
  applyTahsilatDelta,
  applyCekTahsilDelta,
} from "@/lib/finans-ozet";
import { kurus } from "@/lib/num";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

async function yetkili() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isFinance(user.username)) return null;
  return user;
}

export async function GET(req: Request) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const tur = url.searchParams.get("tur");
  const durum = url.searchParams.get("durum");
  let records = await listCekSenet();
  if (tur === "alinan" || tur === "verilen") {
    records = records.filter((c) => c.tur === tur);
  }
  if (durum) records = records.filter((c) => c.durum === durum);
  return NextResponse.json({ ok: true, records });
}

// Yeni çek/senet. Alınan kayıt otomatik Tahsilat üretir (cari düşer, kasa
// değişmez — kasa girişi tahsil geçişinde).
export async function POST(req: Request) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as Record<string, unknown> | null;
  if (!body) {
    return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  }
  const tutar = kurus(Number(body.tutar) || 0);
  const vade = String(body.vade || "");
  const tur = body.tur === "verilen" ? "verilen" : "alinan";
  const kind = body.kind === "senet" ? "senet" : "cek";
  if (tutar <= 0 || !/^\d{4}-\d{2}-\d{2}$/.test(vade)) {
    return NextResponse.json(
      { ok: false, error: "Tutar ve vade gerekli." },
      { status: 400 }
    );
  }
  const customerName = String(body.customerName || "").trim().slice(0, 200);
  const supplier = String(body.supplier || "").trim().slice(0, 200);
  if (tur === "alinan" && !customerName) {
    return NextResponse.json(
      { ok: false, error: "Alınan çekte müşteri adı gerekli." },
      { status: 400 }
    );
  }
  if (tur === "verilen" && !supplier) {
    return NextResponse.json(
      { ok: false, error: "Verilen çekte alıcı (tedarikçi) gerekli." },
      { status: 400 }
    );
  }

  const now = new Date();
  const cs: CekSenet = {
    id: newCekSenetId(now),
    createdAt: now.toISOString(),
    createdBy: user.name,
    tur,
    kind,
    branch: body.branch === "istanbul" ? "istanbul" : "ankara",
    banka: String(body.banka || "").trim().slice(0, 100) || undefined,
    bankaSube: String(body.bankaSube || "").trim().slice(0, 100) || undefined,
    hesapNo: String(body.hesapNo || "").trim().slice(0, 60) || undefined,
    belgeNo: String(body.belgeNo || "").trim().slice(0, 60) || undefined,
    cekSahibi: String(body.cekSahibi || "").trim().slice(0, 200) || undefined,
    tutar,
    vade,
    customerId: String(body.customerId || "").slice(0, 40) || undefined,
    customerName: customerName || undefined,
    supplier: supplier || undefined,
    durum: "portfoyde",
    history: [
      {
        durum: "portfoyde",
        date: istanbulDateKey(now),
        by: user.name,
        note: tur === "alinan" ? "Çek/senet alındı" : "Çek/senet verildi",
      },
    ],
    note: String(body.note || "").trim().slice(0, 300) || undefined,
    kaynak: "panel",
  };

  // Alınan çek cariyi hemen düşürür
  if (tur === "alinan") {
    const t: Tahsilat = {
      id: newTahsilatId(now),
      dateKey: istanbulDateKey(now),
      createdAt: now.toISOString(),
      createdBy: user.name,
      branch: cs.branch,
      customerId: cs.customerId,
      customerName: customerName,
      amount: tutar,
      currency: "TL",
      method: kind,
      cekSenetId: cs.id,
      note: `${kind === "cek" ? "Çek" : "Senet"} — vade ${vade}`,
      kaynak: "panel",
    };
    if (await saveTahsilat(t)) {
      cs.tahsilatId = t.id;
      await applyTahsilatDelta(t, 1).catch(() => {});
    }
  }

  const stored = await saveCekSenet(cs);
  if (!stored) {
    return NextResponse.json(
      { ok: false, error: "Kalıcı depolama yapılandırılmadı." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, ceksenet: cs });
}

// Durum geçişi: { id, durum, date?, note?, ciroTarget? }
export async function PATCH(req: Request) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as Record<string, unknown> | null;
  const id = String(body?.id || "");
  const hedef = String(body?.durum || "") as CekSenetDurum;
  if (!/^CS-[\dA-Za-z-]+$/.test(id)) {
    return NextResponse.json({ ok: false, error: "Geçersiz kayıt." }, { status: 400 });
  }
  const cs = await getCekSenet(id);
  if (!cs) {
    return NextResponse.json({ ok: false, error: "Kayıt bulunamadı." }, { status: 404 });
  }
  if (!allowedTransitions(cs.tur, cs.durum).includes(hedef)) {
    return NextResponse.json(
      { ok: false, error: `"${cs.durum}" durumundan "${hedef}" geçişi yapılamaz.` },
      { status: 400 }
    );
  }
  const dateRaw = String(body?.date || "");
  const date = /^\d{4}-\d{2}-\d{2}$/.test(dateRaw) ? dateRaw : istanbulDateKey();
  const note = String(body?.note || "").trim().slice(0, 300) || undefined;

  if (hedef === "ciro") {
    const target = String(body?.ciroTarget || "").trim().slice(0, 200);
    if (!target) {
      return NextResponse.json(
        { ok: false, error: "Ciro edilen tedarikçi adı gerekli." },
        { status: 400 }
      );
    }
    cs.ciroTarget = target;
    cs.ciroDate = date;
    // Ciro kasaya dokunmaz — çek el değiştirir, tedarikçi borcu kapanır.
  }
  if (hedef === "tahsil") {
    cs.tahsilDate = date;
    // Kasa banka girişi tahsil tarihine yazılır.
    await applyCekTahsilDelta({ ...cs, tahsilDate: date }, 1).catch(() => {});
  }
  if (hedef === "karsiliksiz" || (hedef === "iade" && cs.tur === "alinan")) {
    // Müşterinin borcu geri doğar: girişte oluşan Tahsilat kaydı kaldırılır.
    if (cs.tahsilatId) {
      const girisAyi = ayKey(cs.history[0]?.date || cs.createdAt.slice(0, 10));
      const t = await getTahsilat(girisAyi, cs.tahsilatId);
      if (t) {
        await deleteTahsilat(girisAyi, cs.tahsilatId);
        await applyTahsilatDelta(t, -1).catch(() => {});
      }
    }
  }

  cs.durum = hedef;
  cs.history.push({ durum: hedef, date, by: user.name, note });
  await saveCekSenet(cs);
  return NextResponse.json({ ok: true, ceksenet: cs });
}
