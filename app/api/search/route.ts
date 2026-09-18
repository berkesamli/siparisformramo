import { NextRequest, NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { readOrderIndex, STATUS_LABELS, PAYMENT_LABELS, type OrderIndexEntry } from "@/lib/orders";
import { readRetailIndex, type RetailIndexEntry } from "@/lib/retail-orders";
import { listCustomers, customerTitle, bolgeler, musteriBolgesi } from "@/lib/customers";
import { listRetailCustomers } from "@/lib/retail-customers";
import { getStockData } from "@/lib/stock-store";
import { searchStock, toBoy } from "@/lib/stock-search";
import { eslesir, sikistir } from "@/lib/search-norm";
import { FRAME_PROFILES } from "@/data/catalog";
import { TECHNICAL_PRODUCTS } from "@/data/technical";
import { memo } from "@/lib/server-cache";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 20;

// Genel arama (Ctrl+K). Müşteriler yalnızca ürün/stok arar; çalışanlar
// sipariş, perakende sipariş ve müşteri defterlerinde de arar. Listeler
// 45 sn süreç içi önbellekte tutulur — her tuş vuruşu Blob'a gitmez.

interface Hit {
  kind: "order" | "retail" | "customer" | "retailCustomer" | "stock" | "catalog" | "technical";
  title: string;
  sub?: string;
  href: string;
  meta?: string;
  metaKind?: "ok" | "warn" | "err" | "info" | "brand";
}

const nf = (n: number) => n.toLocaleString("tr-TR");
const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });

const TOPTAN_KIND: Record<string, Hit["metaKind"]> = {
  olusturuldu: "warn",
  hazirlaniyor: "info",
  tamamlandi: "ok",
  iptal: "err",
};
const PERAKENDE_KIND: Record<string, Hit["metaKind"]> = {
  Beklemede: "warn",
  "Hazırlanıyor": "info",
  "Hazır": "ok",
  "Teslim Edildi": undefined,
  "İptal": "err",
};

export async function GET(req: NextRequest) {
  const user = await getSessionUser();
  if (!user) {
    return NextResponse.json({ ok: false, error: "Giriş gerekli." }, { status: 401 });
  }
  const q = (req.nextUrl.searchParams.get("q") || "").trim().slice(0, 80);
  if (q.length < 2) return NextResponse.json({ ok: true, hits: [] });
  const qs = sikistir(q);
  const staff = user.role === "staff";

  const hits: Hit[] = [];

  // ---- Ürün & stok (herkes) ----
  try {
    const stok = await memo("search:stok", 45_000, () => getStockData());
    for (const m of searchStock(stok.items, q, 0.74, 5)) {
      const ank = toBoy(m.item.ankaraMt);
      const ist = toBoy(m.item.istanbulMt);
      hits.push({
        kind: "stock",
        title: m.item.code,
        sub: `Ankara ${nf(ank)} boy · İstanbul ${nf(ist)} boy`,
        href: `/portal?q=${encodeURIComponent(m.item.code)}`,
        meta: ank + ist > 0 ? "Stokta" : "Yok",
        metaKind: ank + ist > 0 ? "ok" : "err",
      });
    }
  } catch { /* stok okunamazsa diğer sonuçlar yeter */ }

  for (const p of FRAME_PROFILES) {
    if (hits.filter((h) => h.kind === "catalog").length >= 5) break;
    if (sikistir(p.code).includes(qs) || sikistir(p.series + p.code).includes(qs)) {
      hits.push({
        kind: "catalog",
        title: p.code,
        sub: `${p.series} serisi · koli ${p.koliAdet} adet / ${p.koliMetraj} mt`,
        href: `/portal/fiyat-listesi?q=${encodeURIComponent(p.code)}`,
        meta: p.stok === "var" ? "Var" : p.stok === "az" ? "Az" : "Yok",
        metaKind: p.stok === "var" ? "ok" : p.stok === "az" ? "warn" : "err",
      });
    }
  }
  for (const t of TECHNICAL_PRODUCTS) {
    if (hits.filter((h) => h.kind === "technical").length >= 5) break;
    if (eslesir(q, t.code, t.name, t.category)) {
      hits.push({
        kind: "technical",
        title: t.name,
        sub: `${t.category} · ${t.code}`,
        href: `/portal/fiyat-listesi?q=${encodeURIComponent(t.code)}&liste=teknik`,
        meta: t.priceEUR ? `€${t.priceEUR}` : t.priceTL ? tl(t.priceTL) : undefined,
        metaKind: "brand",
      });
    }
  }

  if (staff) {
    const [tIdx, pIdx, musteriler, pMusteriler] = await Promise.all([
      memo("search:toptan-idx", 45_000, () => readOrderIndex(6)).catch(() => [] as OrderIndexEntry[]),
      memo("search:perakende-idx", 45_000, () => readRetailIndex(6)).catch(() => [] as RetailIndexEntry[]),
      memo("search:musteriler", 45_000, () => listCustomers()).catch(() => []),
      memo("search:perakende-musteriler", 45_000, () => listRetailCustomers()).catch(() => []),
    ]);

    let n = 0;
    for (const o of tIdx) {
      if (n >= 6) break;
      if (eslesir(q, o.customer, o.orderId, o.employee, o.note)) {
        n++;
        hits.push({
          kind: "order",
          title: `${o.orderId} · ${o.customer || "—"}`,
          sub: `${o.dateKey} · ${o.employee} · ${tl(o.net)}${o.payment ? " · " + PAYMENT_LABELS[o.payment] : ""}`,
          href: `/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
          meta: STATUS_LABELS[o.status],
          metaKind: TOPTAN_KIND[o.status],
        });
      }
    }
    n = 0;
    for (const o of pIdx) {
      if (n >= 6) break;
      if (eslesir(q, o.customerName, o.orderId, o.customerPhone, o.employee)) {
        n++;
        hits.push({
          kind: "retail",
          title: `${o.orderId} · ${o.customerName || "—"}`,
          sub: `${o.dateKey} · ${o.adet} kalem · ${tl(o.total)}${o.deliveryDate ? " · teslim " + o.deliveryDate : ""}`,
          href: `/panel/perakende/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
          meta: o.status,
          metaKind: PERAKENDE_KIND[o.status],
        });
      }
    }
    n = 0;
    for (const c of musteriler) {
      if (n >= 5) break;
      const ad = customerTitle(c);
      if (eslesir(q, ad, c.company, c.firstName, c.lastName, c.phone, c.city, c.email)) {
        n++;
        hits.push({
          kind: "customer",
          title: ad,
          sub: [c.city, c.phone, `${bolgeler()[musteriBolgesi(c)].label} müşterisi`].filter(Boolean).join(" · "),
          href: `/musteriler/kart?id=${encodeURIComponent(c.id)}`,
          meta: c.iskontoPct ? `%${c.iskontoPct} isk.` : undefined,
          metaKind: "brand",
        });
      }
    }
    n = 0;
    for (const c of pMusteriler) {
      if (n >= 4) break;
      if (eslesir(q, c.name, c.phone, c.email)) {
        n++;
        hits.push({
          kind: "retailCustomer",
          title: c.name,
          sub: [c.phone, c.email].filter(Boolean).join(" · "),
          href: `/panel/perakende/musteriler?q=${encodeURIComponent(c.phone || c.name)}`,
        });
      }
    }
  }

  return NextResponse.json({ ok: true, hits });
}
