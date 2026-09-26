import test from "node:test";
import assert from "node:assert/strict";
import { notBloklari, notCoz, varyantAdiCoz, ikasSiparisKalemleri, sayi } from "@/lib/ikas/not";
import type { IkasOrder } from "@/lib/ikas/client";

// Ekran görüntüsündeki iki "Cerceve Hesaplayici" notu (tek satıra sıkıştırılmış hâliyle)
const NOT_1 = `CERCEVE SIPARIS DETAYI (OZL-mugu915y-5303) Urun: Özel Çerçeve: 1266-02 – Eser 100×70 cm (Dış ≈104×74 cm) – Kırılmayan Mat Cam | Adet: 2 [URUN] • SKU (cerceve profili): 1266-02 | Adet: 2 | Profil genisligi: 20 mm | Icerik: Resim / Fotoğraf Baskısı • Yon: Yatay [OLCULER (mm)]
• Sanat eseri: 1000 × 700 mm (musterinin verdigi eser olcusu)
• Kesim: 1002 × 702 mm (cerceve falz / ic kesim olcusu; ic acıklık + 2 mm montaj payi)
• Cerceve ic olcusu (paspartu dahil): 1000 × 700 mm (paspartu yok, eser olcusuyla ayni)
• Tahmini dis olcu (cerceve profili dahil): ~1040 × 740 mm + 2mm pay: Eklendi [PASPARTU] • Paspartu: Yok [CAM] • Cam: Kırılmayan Mat Cam | Cam olcusu: cerceve ic acikligina gore (1000 × 700 mm) [FIYAT]
• Fiyat: Cerceve 1.916,60 TL • Paspartu 0,00 TL • Cam 2.800,00 TL • TOPLAM: 4.716,60 TL
• SUNUCU DOGRULAMASI: metre fiyati 259,00 TL/m (sayfa), tutar bu fiyatla yeniden hesaplandi`;

const NOT_2 = `CERCEVE SIPARIS DETAYI (OZL-mugu66is-6053) Urun: Özel Çerçeve: 1266-02 – Eser 50×70 cm (Dış ≈54×74 cm) – Kırılmayan Mat Cam | Adet: 3 [URUN] • SKU (cerceve profili): 1266-02 | Adet: 3 | Profil genisligi: 20 mm | Icerik: Diploma / Belge • Yon: Dikey [OLCULER (mm)] • Sanat eseri: 500 × 700 mm (musterinin verdigi eser olcusu) • Kesim: 502 × 702 mm (cerceve falz / ic kesim olcusu; ic acıklık + 2 mm montaj payi) • Cerceve ic olcusu (paspartu dahil): 500 × 700 mm (paspartu yok, eser olcusuyla ayni) • Tahmini dis olcu (cerceve profili dahil): ~540 × 740 mm + 2mm pay: Eklendi [PASPARTU] • Paspartu: Yok [CAM] • Cam: Kırılmayan Mat Cam | Cam olcusu: cerceve ic acikligina gore (500 × 700 mm) [FIYAT] • Fiyat: Cerceve 2.097,90 TL • Paspartu 0,00 TL • Cam 2.100,00 TL • TOPLAM: 4.197,90 TL • SUNUCU DOGRULAMASI: metre fiyati 259,00 TL/m (sayfa), tutar bu fiyatla yeniden hesaplandi`;

test("sayi: Türkçe biçimler", () => {
  assert.equal(sayi("1.916,60"), 1916.6);
  assert.equal(sayi("4.197,90"), 4197.9);
  assert.equal(sayi("1000"), 1000);
  assert.equal(sayi("0,00"), 0);
  assert.equal(sayi("259,00"), 259);
  assert.equal(sayi("1.040"), 1040);
});

test("notBloklari: iki blok, OZL kimlikleri", () => {
  const b = notBloklari(NOT_1 + "\n\n" + NOT_2);
  assert.equal(b.length, 2);
  assert.equal(b[0].ozl, "OZL-MUGU915Y-5303");
  assert.equal(b[1].ozl, "OZL-MUGU66IS-6053");
});

test("notCoz: 100×70 yatay, paspartu yok, kırılmayan mat cam", () => {
  const c = notCoz(NOT_1)!;
  assert.ok(c);
  assert.equal(c.sku, "1266-02");
  assert.equal(c.adet, 2);
  assert.equal(c.profilMm, 20);
  assert.equal(c.icerik, "Resim / Fotoğraf Baskısı");
  assert.equal(c.yon, "Yatay");
  assert.deepEqual(c.eserMm, { w: 1000, h: 700 });
  assert.deepEqual(c.kesimMm, { w: 1002, h: 702 });
  assert.deepEqual(c.icMm, { w: 1000, h: 700 });
  assert.deepEqual(c.disMm, { w: 1040, h: 740 });
  assert.equal(c.pay, true);
  assert.equal(c.paspartu, "Yok");
  assert.equal(c.cam, "Kırılmayan Mat Cam");
  assert.deepEqual(c.fiyat, { cerceve: 1916.6, paspartu: 0, cam: 2800, toplam: 4716.6 });
  assert.equal(c.metreFiyat, 259);
});

test("notCoz: 50×70 dikey diploma", () => {
  const c = notCoz(NOT_2)!;
  assert.equal(c.sku, "1266-02");
  assert.equal(c.adet, 3);
  assert.equal(c.icerik, "Diploma / Belge");
  assert.equal(c.yon, "Dikey");
  assert.deepEqual(c.eserMm, { w: 500, h: 700 });
  assert.deepEqual(c.disMm, { w: 540, h: 740 });
  assert.equal(c.fiyat?.toplam, 4197.9);
});

test("notCoz: paspartulu kalem — kenar mm ve iç ölçüden çıkarım", () => {
  const metin = `CERCEVE SIPARIS DETAYI (OZL-abc-1) • SKU (cerceve profili): GB139-1211T | Adet: 1 | Icerik: Belirtilmedi • Yon: Dikey
  • Sanat eseri: 300 × 400 mm • Cerceve ic olcusu (paspartu dahil): 400 × 500 mm • Tahmini dis olcu: ~440 × 540 mm + 2mm pay: Eklenmedi
  [PASPARTU] • Paspartu: Düz Karton (Beyaz) [CAM] • Cam: Düz Cam [FIYAT] • Fiyat: Cerceve 500,00 TL • Paspartu 300,00 TL • Cam 200,00 TL • TOPLAM: 1.000,00 TL`;
  const c = notCoz(metin)!;
  assert.equal(c.sku, "GB139-1211T");
  assert.equal(c.icerik, undefined);
  assert.equal(c.pay, false);
  assert.equal(c.paspartu, "Düz Karton (Beyaz)");
  assert.deepEqual(c.paspartuKenarMm, { ust: 50, alt: 50, sag: 50, sol: 50 });
  const metin2 = metin.replace("Paspartu: Düz Karton (Beyaz)", "Paspartu: Kadife kenar 60 mm");
  assert.deepEqual(notCoz(metin2)!.paspartuKenarMm, { ust: 60, sag: 60, alt: 60, sol: 60 });
});

test("varyantAdiCoz: not yokken varyant adından", () => {
  const c = varyantAdiCoz("Özel Çerçeve: 1266-02 – Eser 50×70 cm (Dış ≈54×74 cm) – Kırılmayan Mat Cam", "OZL-mugu66is-6053")!;
  assert.equal(c.sku, "1266-02");
  assert.deepEqual(c.eserMm, { w: 500, h: 700 });
  assert.deepEqual(c.disMm, { w: 540, h: 740 });
  assert.equal(c.cam, "Kırılmayan Mat Cam");
  assert.equal(c.kaynak, "varyant");
  assert.equal(varyantAdiCoz("Agraf 1000 adet"), null);
});

const siparis = (note: string, satirlar: { sku: string; name: string; qty: number; fiyat: number }[]): IkasOrder => ({
  id: "o1", orderNumber: "1463", status: "CREATED", note,
  orderLineItems: satirlar.map((s, i) => ({ id: "l" + i, quantity: s.qty, price: s.fiyat, finalPrice: s.fiyat, variant: { name: s.name, sku: s.sku }, options: [] })),
});

test("ikasSiparisKalemleri: bloklar OZL ile satırlara eşlenir, föy verisi çıkar", () => {
  const o = siparis(NOT_2 + "\n" + NOT_1, [
    { sku: "OZL-mugu915y-5303", name: "Özel Çerçeve: 1266-02 – Eser 100×70 cm (Dış ≈104×74 cm) – Kırılmayan Mat Cam", qty: 2, fiyat: 4716 },
    { sku: "OZL-mugu66is-6053", name: "Özel Çerçeve: 1266-02 – Eser 50×70 cm (Dış ≈54×74 cm) – Kırılmayan Mat Cam", qty: 3, fiyat: 4197 },
  ]);
  const r = ikasSiparisKalemleri(o);
  assert.equal(r.eksik.length, 0);
  assert.equal(r.kalemler.length, 2);
  assert.equal(r.kalemler[0].adet, 2);
  assert.deepEqual(r.kalemler[0].eserMm, { w: 1000, h: 700 });
  assert.equal(r.kalemler[0].yon, "Yatay");
  assert.equal(r.kalemler[1].adet, 3);
  assert.equal(r.kalemler[1].icerik, "Diploma / Belge");
  assert.equal(r.kalemler[1].retail?.artWidth, 500);
  assert.equal(r.kalemler[1].retail?.artWidthUnit, "mm");
  assert.equal(r.kalemler[1].retail?.glassType, "Kırılmayan Mat Cam");
  assert.equal(r.kalemler[1].retail?.matType, "Paspartu Yok");
  assert.equal(r.kalemler[1].retail?.itemTotal, 1399.3);
  assert.equal(r.kalemler[1].fiyat, 4197.9);
  assert.match(r.kalemler[1].ozet, /Eser 50×70 cm/);
});

test("ikasSiparisKalemleri: not yoksa varyant adı, o da yoksa eksik", () => {
  const o = siparis("", [
    { sku: "OZL-x-1", name: "Özel Çerçeve: KS3420-BLACK – Eser 30×40 cm (Dış ≈36×46 cm) – Düz Cam", qty: 1, fiyat: 900 },
    { sku: "AGR-1", name: "Agraf 1000 adet", qty: 2, fiyat: 100 },
  ]);
  const r = ikasSiparisKalemleri(o);
  assert.equal(r.kalemler[0].sku, "KS3420-BLACK");
  assert.deepEqual(r.kalemler[0].eserMm, { w: 300, h: 400 });
  assert.equal(r.kalemler[0].retail?.glassType, "Düz Cam");
  assert.equal(r.kalemler[1].sku, "AGR-1");
  assert.equal(r.eksik.length, 1);
});
