import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { isFinance } from "@/data/users";
import { listTahsilatByMonths } from "@/lib/tahsilat";
import { listGiderByMonths } from "@/lib/gider";
import { listCekSenet } from "@/lib/ceksenet";
import { istanbulDateKey } from "@/lib/orders";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Kasa raporu — tarih aralığındaki tahsilat + gider + çek tahsil hareketleri.
// Excel'deki aylık kasa günlüğünün karşılığı; yalnızca aralıktaki ay önekleri
// listelenir, tüm geçmiş taranmaz.

function monthsBetween(bas: string, son: string): string[] {
  const out: string[] = [];
  let [y, m] = bas.slice(0, 7).split("-").map(Number);
  const sonAy = son.slice(0, 7);
  for (let i = 0; i < 24; i++) {
    const ay = `${y}-${String(m).padStart(2, "0")}`;
    out.push(ay);
    if (ay === sonAy) break;
    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
  }
  return out;
}

export interface KasaSatir {
  dateKey: string;
  yon: "G" | "C";
  tip: "tahsilat" | "gider" | "cek-tahsil";
  taraf: string; // müşteri / tedarikçi / kategori
  aciklama: string;
  kanal: "nakit" | "banka" | "portfoy";
  currency: string;
  amount: number;
  branch: string;
  kaydeden: string;
}

export async function GET(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff" || !isFinance(user.username)) {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const url = new URL(req.url);
  const bugun = istanbulDateKey();
  const varsayilanBas = bugun.slice(0, 7) + "-01";
  const bas = /^\d{4}-\d{2}-\d{2}$/.test(url.searchParams.get("bas") || "")
    ? url.searchParams.get("bas")!
    : varsayilanBas;
  const son = /^\d{4}-\d{2}-\d{2}$/.test(url.searchParams.get("son") || "")
    ? url.searchParams.get("son")!
    : bugun;
  const sube = url.searchParams.get("sube");

  const months = monthsBetween(bas, son);
  const [tahsilatlar, giderler, cekler] = await Promise.all([
    listTahsilatByMonths(months),
    listGiderByMonths(months),
    listCekSenet(),
  ]);

  const rows: KasaSatir[] = [];

  for (const t of tahsilatlar) {
    if (t.dateKey < bas || t.dateKey > son) continue;
    if (sube && t.branch !== sube) continue;
    const kanal =
      t.method === "nakit"
        ? "nakit"
        : t.method === "havale" || t.method === "krediKarti"
          ? "banka"
          : "portfoy"; // çek/senet — kasada değil, bilgi satırı
    rows.push({
      dateKey: t.dateKey,
      yon: "G",
      tip: "tahsilat",
      taraf: t.customerName,
      aciklama: t.note || "",
      kanal,
      currency: t.currency,
      amount: t.amount,
      branch: t.branch,
      kaydeden: t.tahsilEden || t.createdBy,
    });
  }
  for (const g of giderler) {
    if (g.dateKey < bas || g.dateKey > son) continue;
    if (sube && g.branch !== sube) continue;
    rows.push({
      dateKey: g.dateKey,
      yon: "C",
      tip: "gider",
      taraf: g.supplier || g.category,
      aciklama: [g.category, g.description].filter(Boolean).join(" — "),
      kanal: g.method === "nakit" ? "nakit" : g.method === "cek" ? "portfoy" : "banka",
      currency: g.currency,
      amount: g.amount,
      branch: g.branch,
      kaydeden: g.createdBy,
    });
  }
  // Çek tahsilleri: kasa banka girişi tahsil tarihinde
  for (const c of cekler) {
    if (c.durum !== "tahsil" || !c.tahsilDate) continue;
    if (c.tahsilDate < bas || c.tahsilDate > son) continue;
    if (sube && c.branch !== sube) continue;
    rows.push({
      dateKey: c.tahsilDate,
      yon: "G",
      tip: "cek-tahsil",
      taraf: c.customerName || "-",
      aciklama: `${c.kind === "cek" ? "Çek" : "Senet"} tahsili — vade ${c.vade}`,
      kanal: "banka",
      currency: "TL",
      amount: c.tutar,
      branch: c.branch,
      kaydeden: c.createdBy,
    });
  }

  rows.sort((a, b) => b.dateKey.localeCompare(a.dateKey));

  const r2 = (n: number) => Math.round(n * 100) / 100;
  const topla = (f: (r: KasaSatir) => boolean) =>
    r2(rows.filter(f).reduce((s, r) => s + r.amount, 0));

  const ozet = {
    girisNakit: topla((r) => r.yon === "G" && r.kanal === "nakit" && r.currency === "TL"),
    girisBanka: topla((r) => r.yon === "G" && r.kanal === "banka" && r.currency === "TL"),
    cikisNakit: topla((r) => r.yon === "C" && r.kanal === "nakit" && r.currency === "TL"),
    cikisBanka: topla((r) => r.yon === "C" && r.kanal === "banka" && r.currency === "TL"),
    portfoyGiris: topla((r) => r.yon === "G" && r.kanal === "portfoy"),
    dovizUsd: topla((r) => r.currency === "USD" && r.yon === "G") -
      topla((r) => r.currency === "USD" && r.yon === "C"),
    dovizEur: topla((r) => r.currency === "EUR" && r.yon === "G") -
      topla((r) => r.currency === "EUR" && r.yon === "C"),
  };

  return NextResponse.json({ ok: true, bas, son, rows: rows.slice(0, 1500), ozet });
}
