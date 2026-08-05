import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isOwner } from "@/data/users";
import {
  getMaliyetData,
  saveMaliyetData,
  birimMaliyetTL,
  normKod,
  type MaliyetKaydi,
  type AlisBirimi,
} from "@/lib/maliyet";
import { listAllOrders } from "@/lib/orders";
import { findProfile } from "@/data/catalog";
import { kurus } from "@/lib/num";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Alış fiyatları ve kod bazlı kâr analizi — YALNIZCA firma sahipleri.
// (FINANS_AKTIF bayrağından bağımsızdır; bu bölüm ayrıca istendi.)

async function sahip() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isOwner(user.username)) return null;
  return user;
}

// GET  → kayıtlı alış fiyatları; ?analiz=1&ay=YYYY-MM ile satış analizi
export async function GET(req: NextRequest) {
  const user = await sahip();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const data = await getMaliyetData();

  if (req.nextUrl.searchParams.get("analiz") !== "1") {
    return NextResponse.json({ ok: true, data });
  }

  // ---- Satış analizi: toptan sipariş satırları kod bazında toplanır ----
  const ay = req.nextUrl.searchParams.get("ay") || "";
  const orders = (await listAllOrders()).filter((o) =>
    ay ? o.dateKey.startsWith(ay) : true
  );

  interface KodAnaliz {
    code: string;
    metraj: number;
    ciro: number;
    maliyet: number | null; // alış girilmemişse null
    satirSayisi: number;
  }
  const map = new Map<string, KodAnaliz>();

  for (const o of orders) {
    const usd = Number(o.rate) || 0;
    const eur = Number(o.euroRate) || usd;
    for (const l of o.lines) {
      // "4501S-1242" → taban kod; katalogda karşılığı varsa onun kodu
      const ham = String(l.name || "").trim().split(/[\s(]/)[0];
      if (!ham) continue;
      const p = findProfile(ham);
      const kod = p ? normKod(p.code) : normKod(ham.split("-")[0]);
      if (!kod) continue;
      // metraj = satır tutarı / birim TL fiyat (satırlar mt bazlı fiyatlanır)
      const metre =
        l.unitPriceTL > 0 ? l.lineTotal / l.unitPriceTL : 0;
      const cur = map.get(kod) || {
        code: p ? p.code : ham.split("-")[0],
        metraj: 0,
        ciro: 0,
        maliyet: 0 as number | null,
        satirSayisi: 0,
      };
      cur.metraj += metre;
      cur.ciro += l.lineTotal;
      cur.satirSayisi += 1;
      const mk = data.items[kod];
      if (mk && usd > 0 && cur.maliyet != null) {
        cur.maliyet += metre * birimMaliyetTL(mk, data.defaultPct, usd, eur);
      } else if (!mk) {
        cur.maliyet = null; // alışı girilmemiş kod — kâr hesaplanamaz
      }
      map.set(kod, cur);
    }
  }

  const analiz = [...map.values()]
    .map((k) => ({
      code: k.code,
      metraj: Math.round(k.metraj * 100) / 100,
      ciro: kurus(k.ciro),
      maliyet: k.maliyet != null ? kurus(k.maliyet) : null,
      kar: k.maliyet != null ? kurus(k.ciro - k.maliyet) : null,
      marj:
        k.maliyet != null && k.ciro > 0
          ? Math.round(((k.ciro - k.maliyet) / k.ciro) * 1000) / 10
          : null,
      satirSayisi: k.satirSayisi,
    }))
    .sort((a, b) => b.ciro - a.ciro);

  const toplamCiro = kurus(analiz.reduce((s, a) => s + a.ciro, 0));
  const maliyetliler = analiz.filter((a) => a.maliyet != null);
  const toplamMaliyet = kurus(maliyetliler.reduce((s, a) => s + (a.maliyet || 0), 0));
  const maliyetliCiro = kurus(maliyetliler.reduce((s, a) => s + a.ciro, 0));

  return NextResponse.json({
    ok: true,
    data,
    analiz,
    ozet: {
      toplamCiro,
      maliyetliCiro,
      toplamMaliyet,
      toplamKar: kurus(maliyetliCiro - toplamMaliyet),
      kapsam: analiz.length ? Math.round((maliyetliler.length / analiz.length) * 100) : 0,
    },
  });
}

// POST → { defaultPct?, items?: [{code, alis, currency, pct?, note?}], sil?: [code] }
export async function POST(req: NextRequest) {
  const user = await sahip();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as {
    defaultPct?: unknown;
    items?: unknown[];
    sil?: unknown[];
  } | null;
  if (!body) {
    return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  }

  const data = await getMaliyetData();

  if (body.defaultPct !== undefined) {
    const pct = Number(body.defaultPct);
    if (!Number.isFinite(pct) || pct < 0 || pct > 500) {
      return NextResponse.json({ ok: false, error: "Geçersiz yüzde." }, { status: 400 });
    }
    data.defaultPct = Math.round(pct * 100) / 100;
  }

  const hatalar: string[] = [];
  for (const raw of (body.items || []) as Partial<MaliyetKaydi>[]) {
    const code = String(raw.code || "").trim();
    const alis = Number(raw.alis);
    const currency = (raw.currency || "USD") as AlisBirimi;
    if (!code || !Number.isFinite(alis) || alis <= 0 ||
        !["USD", "EUR", "TL"].includes(currency)) {
      hatalar.push(`geçersiz kayıt: ${code || "?"}`);
      continue;
    }
    const pct = raw.pct != null && raw.pct !== ("" as unknown) ? Number(raw.pct) : undefined;
    data.items[normKod(code)] = {
      code,
      alis: Math.round(alis * 10000) / 10000,
      currency,
      pct: pct != null && Number.isFinite(pct) && pct >= 0 ? pct : undefined,
      note: String(raw.note || "").trim().slice(0, 200) || undefined,
      updatedAt: new Date().toISOString(),
      by: user.name,
    };
  }
  for (const c of (body.sil || []) as string[]) {
    delete data.items[normKod(String(c))];
  }

  data.updatedAt = new Date().toISOString();
  const stored = await saveMaliyetData(data);
  if (!stored) {
    return NextResponse.json(
      { ok: false, error: "Kalıcı depolama yapılandırılmadı." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, data, hatalar });
}
