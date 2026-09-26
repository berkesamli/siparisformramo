import test, { before } from "node:test";
import assert from "node:assert/strict";
import { newDb } from "pg-mem";
import * as db from "@/lib/uretim/db";

before(() => {
  const mem = newDb({ autoCreateForeignKeyIndices: true });
  const { Pool } = mem.adapters.createPg();
  db.setUretimPool(new Pool());
});

test("iş ekle / bul / listele / güncelle — olay günlüğü", async () => {
  const is = await db.isEkle({ kaynak: "elle", baslik: "Ahmet Yılmaz", musteriAd: "Ahmet Yılmaz", sube: "ankara", kalemler: [{ sku: "1266-02", adet: 2, ozet: "Eser 50×70" }], tutar: 1200 }, "test");
  assert.equal(is.durum, "yeni");
  assert.equal(is.planTarih, null);
  assert.equal(is.adet, 2);
  assert.equal(is.tutar, 1200);
  const plansiz = await db.isler({ from: "2026-09-21", to: "2026-09-27", plansiz: true });
  assert.equal(plansiz.length, 1);
  const g = await db.isGuncelle(is.id, { planTarih: "2026-09-24" }, "test");
  assert.equal(g?.planTarih, "2026-09-24");
  assert.equal(g?.durum, "planlandi", "takvime alınınca yeni → planlandı");
  assert.equal(g?.planSira, 1);
  const g2 = await db.isGuncelle(is.id, { durum: "uretimde", sube: "istanbul" }, "test");
  assert.equal(g2?.durum, "uretimde");
  assert.equal(g2?.sube, "istanbul");
  const olaylar = await db.olaylar(is.id);
  const turler = olaylar.map((o) => o.tur);
  assert.ok(turler.includes("olustur") && turler.includes("tasi") && turler.includes("durum") && turler.includes("sube"), turler.join(","));
  const hafta = await db.isler({ from: "2026-09-21", to: "2026-09-27" });
  assert.equal(hafta.length, 1);
  const bos = await db.isler({ from: "2026-10-01", to: "2026-10-07" });
  assert.equal(bos.length, 0);
  const tamam = await db.isGuncelle(is.id, { durum: "tamamlandi" }, "test");
  assert.ok(tamam?.tamamlandiAt);
  const q = await db.isler({ from: "2026-09-21", to: "2026-09-27", q: "ahmet" });
  assert.equal(q.length, 1);
});

test("kaynaktanYaz: yeni → var olan (eski kaynak) → yeni kaynak durumu", async () => {
  const kayit = {
    kaynak: "perakende" as const, kaynakRef: "PRK-2026-001", kaynakKey: "2026-09-20", kaynakAt: "2026-09-20T10:00:00.000Z",
    kaynakDurum: "Beklemede", durum: "planlandi" as const, musteriAd: "Ayşe", musteriTel: "0532", teslimTarih: "2026-09-30",
    kalemler: [{ sku: "GB139", adet: 1, ozet: "x" }], tutar: 500,
  };
  const a = await db.kaynaktanYaz(kayit, "senk");
  assert.equal(a.yeni, true);
  assert.equal(a.is.planTarih, "2026-09-30", "plan = teslim tarihi");
  assert.equal(a.is.durum, "planlandi");
  // Takvimde elle ilerletildi
  await db.isGuncelle(a.is.id, { durum: "uretimde", planTarih: "2026-09-28" }, "eren");
  // Kaynak daha eski → dokunma
  const b = await db.kaynaktanYaz(kayit, "senk");
  assert.equal(b.yeni, false); assert.equal(b.guncellendi, false);
  // Kaynak yeni ama durum aynı → durum korunur, plan tarihi korunur, kalemler tazelenir
  const c = await db.kaynaktanYaz({ ...kayit, kaynakAt: "2026-09-21T10:00:00.000Z", tutar: 600 }, "senk");
  assert.equal(c.guncellendi, true);
  assert.equal(c.is.durum, "uretimde");
  assert.equal(c.is.planTarih, "2026-09-28");
  assert.equal(c.is.tutar, 600);
  // Kaynak durumu değişti → eşlenir
  const d = await db.kaynaktanYaz({ ...kayit, kaynakAt: "2026-09-22T10:00:00.000Z", kaynakDurum: "Teslim Edildi", durum: "teslim" }, "senk");
  assert.equal(d.is.durum, "teslim");
  // Teslim tarihi: kullanıcı değiştirmemişse kaynaktan gelir
  const e = await db.kaynaktanYaz({ ...kayit, kaynakAt: "2026-09-23T10:00:00.000Z", kaynakDurum: "Teslim Edildi", durum: "teslim", teslimTarih: "2026-10-02" }, "senk");
  assert.equal(e.is.teslimTarih, "2026-10-02");
  assert.equal(e.is.planTarih, "2026-09-28", "plan tarihi kaynaktan ezilmez");
  const ayni = await db.isBulKaynak("perakende", "PRK-2026-001");
  assert.equal(ayni?.id, a.is.id);
});

test("gün notları", async () => {
  const n = await db.notEkle({ tarih: "2026-09-25", sube: "ankara", metin: "Kesim makinesi bakım" }, "eren");
  assert.equal(n.tamam, false);
  const g = await db.notGuncelle(n.id, { tamam: true });
  assert.equal(g?.tamam, true);
  const liste = await db.notlar("2026-09-21", "2026-09-27");
  assert.equal(liste.length, 1);
  assert.equal(await db.notSil(n.id), true);
  assert.equal((await db.notlar("2026-09-21", "2026-09-27")).length, 0);
});

test("teklif: sıralı numara, listeleme, güncelleme", async () => {
  const t1 = await db.teklifEkle({ musteriAd: "Mehmet", kalemler: [], brut: 1000, toplam: 900, iskonto: 100, takipTarih: "2026-09-26" }, "berke");
  const t2 = await db.teklifEkle({ musteriAd: "Zeynep", kalemler: [], brut: 500, toplam: 500 }, "berke");
  assert.match(t1.no, /^TKL-\d{4}-001$/);
  assert.match(t2.no, /^TKL-\d{4}-002$/);
  const acik = await db.teklifler({ durum: "acik", takipFrom: "2026-09-21", takipTo: "2026-09-27" });
  assert.equal(acik.length, 1);
  const g = await db.teklifGuncelle(t1.id, { durum: "kabul", siparisRef: "PRK-2026-009" });
  assert.equal(g?.durum, "kabul");
  assert.equal(g?.siparisRef, "PRK-2026-009");
  const q = await db.teklifler({ q: "zeyn" });
  assert.equal(q.length, 1);
});

test("özet ve rapor", async () => {
  await db.isEkle({ kaynak: "online", kaynakRef: "1463", kaynakKey: "x", musteriAd: "Mert", kalemler: [{ sku: "1266-02", adet: 5, ozet: "" }] }, "t");
  await db.isEkle({ kaynak: "elle", musteriAd: "Geciken", planTarih: "2026-09-20", sube: "istanbul", kalemler: [{ sku: "A", adet: 1, ozet: "" }] }, "t");
  await db.isEkle({ kaynak: "elle", musteriAd: "Bugün", planTarih: "2026-09-26", sube: "ankara", kalemler: [{ sku: "A", adet: 3, ozet: "" }] }, "t");
  const o = await db.ozet("2026-09-26");
  assert.equal(o.bugunPlanli.ankara, 1);
  assert.equal(o.bugunPlanli.adet, 3);
  assert.ok(o.plansiz >= 1);
  assert.equal(o.geciken, 1);
  assert.equal(o.yeniOnline, 1);
  const r = await db.rapor("2026-09-20", "2026-09-30", "2026-09-26");
  assert.ok(r.gunler.find((g) => g.tarih === "2026-09-26")?.ankara === 1);
  assert.equal(r.geciken.length, 1);
  // İstanbul: "Geciken" (09-20) + ilk testin 09-24'e planlanıp tamamlanan işi (tamamlananlar rapora dahil, iptal hariç)
  assert.equal(r.toplam.istanbul, 2);
});
