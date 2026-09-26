import test from "node:test";
import assert from "node:assert/strict";
import { isSiparisKaydi, isFoyEkstra, foyUret } from "@/lib/uretim/foy";
import { ikasSiparisKalemleri } from "@/lib/ikas/not";
import type { UretimIs } from "@/lib/uretim/tur";

const NOT = `CERCEVE SIPARIS DETAYI (OZL-mugu66is-6053) • SKU (cerceve profili): 1266-02 | Adet: 3 | Profil genisligi: 20 mm | Icerik: Diploma / Belge • Yon: Dikey [OLCULER (mm)] • Sanat eseri: 500 × 700 mm • Kesim: 502 × 702 mm • Cerceve ic olcusu (paspartu dahil): 500 × 700 mm • Tahmini dis olcu (cerceve profili dahil): ~540 × 740 mm + 2mm pay: Eklendi [PASPARTU] • Paspartu: Yok [CAM] • Cam: Kırılmayan Mat Cam [FIYAT] • Fiyat: Cerceve 2.097,90 TL • Paspartu 0,00 TL • Cam 2.100,00 TL • TOPLAM: 4.197,90 TL`;

function ornekIs(): UretimIs {
  const { kalemler } = ikasSiparisKalemleri({
    id: "o1", orderNumber: "1463", status: "CREATED", note: NOT,
    orderLineItems: [{ id: "l1", quantity: 3, price: 4197, finalPrice: 4197, variant: { name: "Özel Çerçeve: 1266-02 – Eser 50×70 cm (Dış ≈54×74 cm) – Kırılmayan Mat Cam", sku: "OZL-mugu66is-6053" }, options: [] }],
  });
  const now = new Date().toISOString();
  return {
    id: "uabc", kaynak: "online", kaynakRef: "1463", kaynakKey: "o1", baslik: "Mert Kalfa", musteriAd: "Mert Kalfa", musteriTel: "0505 555 88 17",
    musteriAdres: "Orta Mah. Üniversite Cd. No:27 Tuzla İstanbul", musteriSehir: "İstanbul", sube: "istanbul", subeOneri: "istanbul", subeOneriNeden: "",
    planTarih: null, planSaat: "", planSira: 0, teslimTarih: null, durum: "yeni", kalemler, adet: 3, tutar: 8913, notlar: "", foyYol: null, foyAt: null,
    ekler: [], ham: {}, olusturan: "test", createdAt: now, updatedAt: now, tamamlandiAt: null, kaynakAt: now,
  };
}

test("isSiparisKaydi: online işten föy kaydı ve ekstra", () => {
  const is = ornekIs();
  const o = isSiparisKaydi(is)!;
  assert.ok(o);
  assert.equal(o.items.length, 1);
  assert.equal(o.items[0].artWidth, 500);
  assert.equal(o.orderId, "Online #1463");
  assert.match(o.notes, /TOPLAM 3 ADET/);
  const e = isFoyEkstra(is);
  assert.equal(e.sube, "istanbul");
  assert.equal(e.icerik, "Diploma / Belge");
  assert.equal(e.yon, "Dikey");
  assert.match(String(e.kaynak), /1463/);
});

test("foyUret: pdfkit ile şube rozetli föy üretir (%PDF)", async () => {
  const pdf = await foyUret(ornekIs());
  assert.ok(pdf && pdf.length > 5000, "PDF üretilmeli");
  assert.equal(pdf!.subarray(0, 4).toString(), "%PDF");
});

test("foyUret: ölçüsüz kalemde null", async () => {
  const is = ornekIs();
  is.kalemler = [{ sku: "AGR", adet: 1, ozet: "Agraf" }];
  assert.equal(await foyUret(is), null);
});
