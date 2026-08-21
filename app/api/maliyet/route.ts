import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isMaliyet } from "@/data/users";
import {
  getMaliyetData,
  saveMaliyetData,
  kodIndeksi,
  gecerliKalem,
  birimMaliyetTL,
  normKod,
  newPartiId,
  type PartiKalemi,
  type AlisBirimi,
} from "@/lib/maliyet";
import { listAllOrders } from "@/lib/orders";
import { findProfile } from "@/data/catalog";
import { GLASS_TYPES } from "@/data/glass";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { kurus } from "@/lib/num";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Parti (konteyner) bazlı alış fiyatları ve kâr analizi.
// Erişim: yalnızca MALIYET_USERNAMES (varsayılan berke + özgür).

async function yetkili() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isMaliyet(user.username)) return null;
  return user;
}

export async function GET(req: NextRequest) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const data = await getMaliyetData();

  if (req.nextUrl.searchParams.get("analiz") !== "1") {
    return NextResponse.json({ ok: true, data });
  }

  const ay = req.nextUrl.searchParams.get("ay") || "";
  // İptal edilen siparişler satış/kâr analizine girmez
  const orders = (await listAllOrders()).filter(
    (o) => o.status !== "iptal" && (ay ? o.dateKey.startsWith(ay) : true)
  );
  const indeks = kodIndeksi(data);

  // Satır adı → maliyet anahtarı eşlemesi.
  // Cam/ayna satırları tür adıyla ("Düz Cam"), teknik malzeme ürün adıyla
  // ("NS Karton Kadife") yazılır — çerçevedeki "ilk kelime = kod" kuralı
  // onlarda çalışmaz; ad üzerinden eşlenir. Miktar zaten satırın kendi
  // biriminde gelir: çerçevede mt, camda m², teknikte kutu
  // (miktar = tutar ÷ birim fiyat).
  const camAd = new Map<string, string>();
  for (const g of GLASS_TYPES) camAd.set(normKod(g.name), g.name);
  camAd.set(normKod("Ayna"), "Ayna");
  const teknikAd = new Map<string, string>();
  for (const t of TECHNICAL_PRODUCTS) teknikAd.set(normKod(t.name), t.name);

  interface KodAnaliz {
    code: string;
    metraj: number;
    ciro: number;
    maliyet: number;
    maliyetTam: boolean; // tüm satırların maliyeti hesaplanabildi mi
    durum: "" | "alis-yok" | "yuzde-bekliyor";
    satirSayisi: number;
  }
  const map = new Map<string, KodAnaliz>();

  for (const o of orders) {
    const usd = Number(o.rate) || 0;
    const eur = Number(o.euroRate) || usd;
    for (const l of o.lines) {
      const tamAd = String(l.name || "").trim();
      if (!tamAd) continue;

      // Önce tam adla, eşleşmezse sondaki parantezleri birer birer atarak
      // dene — karton kodu "(107)" ve iskonto "(%10 isk.)" ekleri düşer ama
      // adın parçası olan parantez ("... (6x4cm)") korunur.
      let kod = "";
      let goster = "";
      let aday = tamAd;
      for (;;) {
        const a = normKod(aday);
        if (camAd.has(a)) {
          kod = a;
          goster = camAd.get(a)!;
          break;
        }
        if (teknikAd.has(a)) {
          kod = a;
          goster = teknikAd.get(a)!;
          break;
        }
        const kisalt = aday.replace(/\s*\([^)]*\)\s*$/, "").trim();
        if (kisalt === aday || !kisalt) break;
        aday = kisalt;
      }
      if (!kod) {
        const ham = tamAd.split(/[\s(]/)[0];
        const p = findProfile(ham);
        kod = p ? normKod(p.code) : normKod(ham.split("-")[0]);
        goster = p ? p.code : ham.split("-")[0];
      }
      if (!kod) continue;
      const metre = l.unitPriceTL > 0 ? l.lineTotal / l.unitPriceTL : 0;

      const cur = map.get(kod) || {
        code: goster,
        metraj: 0, ciro: 0, maliyet: 0,
        maliyetTam: true, durum: "" as KodAnaliz["durum"], satirSayisi: 0,
      };
      cur.metraj += metre;
      cur.ciro += l.lineTotal;
      cur.satirSayisi += 1;

      const secim = gecerliKalem(indeks.get(kod), o.dateKey);
      if (!secim) {
        cur.maliyetTam = false;
        cur.durum = "alis-yok";
      } else {
        const bm = usd > 0 ? birimMaliyetTL(secim, usd, eur) : null;
        if (bm == null) {
          cur.maliyetTam = false;
          if (cur.durum !== "alis-yok") cur.durum = "yuzde-bekliyor";
        } else {
          cur.maliyet += metre * bm;
        }
      }
      map.set(kod, cur);
    }
  }

  const analiz = [...map.values()]
    .map((k) => ({
      code: k.code,
      metraj: Math.round(k.metraj * 100) / 100,
      ciro: kurus(k.ciro),
      maliyet: k.maliyetTam ? kurus(k.maliyet) : null,
      kar: k.maliyetTam ? kurus(k.ciro - k.maliyet) : null,
      marj:
        k.maliyetTam && k.ciro > 0
          ? Math.round(((k.ciro - k.maliyet) / k.ciro) * 1000) / 10
          : null,
      durum: k.maliyetTam ? "" : k.durum,
      satirSayisi: k.satirSayisi,
    }))
    .sort((a, b) => b.ciro - a.ciro);

  const tamlar = analiz.filter((a) => a.maliyet != null);
  const maliyetliCiro = kurus(tamlar.reduce((s, a) => s + a.ciro, 0));
  const toplamMaliyet = kurus(tamlar.reduce((s, a) => s + (a.maliyet || 0), 0));

  return NextResponse.json({
    ok: true,
    data,
    analiz,
    ozet: {
      toplamCiro: kurus(analiz.reduce((s, a) => s + a.ciro, 0)),
      maliyetliCiro,
      toplamMaliyet,
      toplamKar: kurus(maliyetliCiro - toplamMaliyet),
      kapsam: analiz.length
        ? Math.round((tamlar.length / analiz.length) * 100)
        : 0,
    },
  });
}

// POST eylemleri:
//   { yeniParti: { ad, tarih, note? } }
//   { partiId, pct }                  → yüzdeyi (sonradan) gir/güncelle
//   { partiId, items: [{code, alis, currency}] }
//   { partiId, sil: [code] }
//   { partiSil: partiId }             → boş/yanlış açılan partiyi kaldır
export async function POST(req: NextRequest) {
  const user = await yetkili();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const body = (await req.json().catch(() => null)) as Record<string, unknown> | null;
  if (!body) {
    return NextResponse.json({ ok: false, error: "Geçersiz istek." }, { status: 400 });
  }

  const data = await getMaliyetData();
  const hatalar: string[] = [];

  if (body.yeniParti && typeof body.yeniParti === "object") {
    const yp = body.yeniParti as { ad?: unknown; tarih?: unknown; note?: unknown };
    const ad = String(yp.ad || "").trim().slice(0, 120);
    const tarih = String(yp.tarih || "");
    if (!ad || !/^\d{4}-\d{2}-\d{2}$/.test(tarih)) {
      return NextResponse.json(
        { ok: false, error: "Parti adı ve geliş tarihi gerekli." },
        { status: 400 }
      );
    }
    data.partiler.push({
      id: newPartiId(),
      ad,
      tarih,
      pct: null,
      note: String(yp.note || "").trim().slice(0, 300) || undefined,
      createdAt: new Date().toISOString(),
      by: user.name,
      items: {},
    });
  }

  const partiId = String(body.partiId || "");
  if (partiId) {
    const parti = data.partiler.find((x) => x.id === partiId);
    if (!parti) {
      return NextResponse.json({ ok: false, error: "Parti bulunamadı." }, { status: 404 });
    }
    if (body.pct !== undefined) {
      if (body.pct === null || body.pct === "") {
        parti.pct = null;
      } else {
        const pct = Number(body.pct);
        if (!Number.isFinite(pct) || pct < 0 || pct > 500) {
          return NextResponse.json({ ok: false, error: "Geçersiz yüzde." }, { status: 400 });
        }
        parti.pct = Math.round(pct * 100) / 100;
      }
    }
    for (const raw of (body.items || []) as Partial<PartiKalemi>[]) {
      const code = String(raw.code || "").trim();
      const alis = Number(raw.alis);
      const currency = (raw.currency || "USD") as AlisBirimi;
      const birim = (["mt", "m2", "kutu"].includes(String(raw.birim))
        ? raw.birim
        : "mt") as PartiKalemi["birim"];
      if (!code || !Number.isFinite(alis) || alis <= 0 ||
          !["USD", "EUR", "TL"].includes(currency)) {
        hatalar.push(`geçersiz kalem: ${code || "?"}`);
        continue;
      }
      parti.items[normKod(code)] = {
        code,
        alis: Math.round(alis * 10000) / 10000,
        currency,
        birim,
      };
    }
    for (const c of (body.sil || []) as string[]) {
      delete parti.items[normKod(String(c))];
    }
  }

  if (body.partiSil) {
    data.partiler = data.partiler.filter((x) => x.id !== String(body.partiSil));
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
