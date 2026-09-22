import { NextRequest, NextResponse } from "next/server";
import { getUsers, isMesajci } from "@/data/users";
import { konusmalar, okunmamisSayisi } from "@/lib/mesaj/db";
import { gmailConfigured, gmailSenk } from "@/lib/mesaj/gmail";
import { kanalDurumu } from "@/lib/mesaj/gonder";
import { taslakHazir } from "@/lib/mesaj/taslak";
import { hataYaniti, mesajKullanici } from "@/lib/mesaj/yetki";
import type { Kanal, KonusmaDurum } from "@/lib/mesaj/tur";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// Gelen kutusu listesi. ?senk=1 → Gmail hesapları da tazelenir (45 sn'de en çok bir kez).
export async function GET(req: NextRequest) {
  const y = await mesajKullanici();
  if ("hata" in y) return y.hata;
  const q = req.nextUrl.searchParams;
  try {
    let senk: Awaited<ReturnType<typeof gmailSenk>> | undefined;
    if (q.get("senk") === "1" && gmailConfigured()) senk = await gmailSenk({ zorla: q.get("zorla") === "1" });
    const liste = await konusmalar({
      kanal: (q.get("kanal") || "") as Kanal | "",
      durum: (q.get("durum") || "") as KonusmaDurum | "",
      atanan: q.get("atanan") || "",
      q: (q.get("q") || "").trim().slice(0, 80),
      limit: Number(q.get("limit")) || 80,
    }, y.user.username);
    const kullanicilar = getUsers().filter((u) => u.role === "staff" && isMesajci(u.username)).map((u) => ({ username: u.username, name: u.name }));
    return NextResponse.json({
      ok: true,
      konusmalar: liste,
      okunmamis: await okunmamisSayisi(),
      kanallar: kanalDurumu(),
      taslak: taslakHazir(),
      kullanicilar,
      senk,
    });
  } catch (e) {
    console.error("Konuşma listesi alınamadı:", e);
    return hataYaniti(e);
  }
}
