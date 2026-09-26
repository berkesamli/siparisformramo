// Üretim Takvimi — istemci ve sunucunun paylaştığı tipler.
//
// "İş" (uretim_is): üretilecek her şey tek tabloda durur; kaynağı ne olursa
// olsun (online/ikas siparişi, mağaza/perakende siparişi, teklif, elle giriş)
// aynı takvimde planlanır. Online siparişler ile mağaza siparişleri ayrı
// KAYNAK olarak tutulur, birbirine karışmaz; takvimde ayrı renk ve ayrı
// filtreyle görünür.
//
// Tarihler: planTarih / teslimTarih / gün notu tarihi "YYYY-MM-DD" (İstanbul
// günü); zaman damgaları ISO metin. Para TL.

import type { RetailItem } from "@/lib/retail-orders";

// ---- Kaynak ----
export type IsKaynak = "online" | "perakende" | "toptan" | "teklif" | "elle";

export const KAYNAK_LABELS: Record<IsKaynak, string> = {
  online: "Online Sipariş",
  perakende: "Mağaza Siparişi",
  toptan: "Toptan Sipariş",
  teklif: "Teklif",
  elle: "Elle Eklenen",
};

/** Kaynak kısa etiketi (kart rozeti). */
export const KAYNAK_KISA: Record<IsKaynak, string> = {
  online: "Online",
  perakende: "Mağaza",
  toptan: "Toptan",
  teklif: "Teklif",
  elle: "Elle",
};

// ---- Durum akışı ----
export type IsDurum = "yeni" | "planlandi" | "uretimde" | "hazir" | "teslim" | "tamamlandi" | "iptal";

export const DURUM_LABELS: Record<IsDurum, string> = {
  yeni: "Yeni",
  planlandi: "Planlandı",
  uretimde: "Üretimde",
  hazir: "Hazır",
  teslim: "Teslim / Kargo",
  tamamlandi: "Tamamlandı",
  iptal: "İptal",
};

/** Akış sırası — "sonraki adım" düğmesi bu sırayı izler (iptal dışarıda). */
export const DURUM_SIRA: IsDurum[] = ["yeni", "planlandi", "uretimde", "hazir", "teslim", "tamamlandi"];

/** Takvimde ve kuyrukta "açık" sayılan durumlar. */
export const ACIK_DURUMLAR: IsDurum[] = ["yeni", "planlandi", "uretimde", "hazir"];

export function sonrakiDurum(d: IsDurum): IsDurum | null {
  const i = DURUM_SIRA.indexOf(d);
  return i >= 0 && i < DURUM_SIRA.length - 1 ? DURUM_SIRA[i + 1] : null;
}

export function isAcik(d: IsDurum): boolean {
  return ACIK_DURUMLAR.includes(d);
}

// ---- Şube ----
export type Sube = "ankara" | "istanbul" | "";

export const SUBE_LABELS: Record<Exclude<Sube, "">, string> = {
  ankara: "Ankara",
  istanbul: "İstanbul",
};

// ---- Kalem ----
/**
 * Üretim kalemi. `retail` doluysa kalem perakende fiyat çekirdeğiyle aynı
 * yapıdadır ve üretim föyü (generateRetailPdf) buradan üretilir. Özet
 * alanlar takvim kartı ve yan panel içindir; föy için retail gerekir.
 */
export interface IsKalem {
  sku: string;          // çerçeve profili kodu (1266-02, GB139-1211T …)
  adet: number;
  ozet: string;         // "Eser 50×70 cm · Kırılmayan Mat Cam · Paspartu yok"
  eserMm?: { w: number; h: number };
  disMm?: { w: number; h: number };  // tahmini dış ölçü
  icerik?: string;      // "Resim / Fotoğraf Baskısı", "Diploma / Belge"…
  yon?: string;         // Yatay / Dikey / Kare
  paspartu?: string;    // "Yok" ya da "Düz Karton 5 cm"
  cam?: string;         // "Kırılmayan Mat Cam"
  pay?: boolean;        // +2 mm kesim payı eklendi mi
  fiyat?: number;       // kalem toplamı (TL)
  retail?: RetailItem;  // föy üretimi için tam kalem
}

export interface IsEk {
  id: string;
  ad: string;
  yol: string;          // Blob yolu (uretim/ek/<isId>/<id>-<ad>)
  tur: "pdf" | "image" | "file";
  boyut: number;
  at: string;
  by: string;
}

export interface UretimIs {
  id: string;
  kaynak: IsKaynak;
  kaynakRef: string;    // PRK-2026-001 · ikas sipariş no (1463) · teklif no · ""
  kaynakKey: string;    // perakende: siparişin dateKey'i · online: ikas sipariş id'si
  baslik: string;
  musteriAd: string;
  musteriTel: string;
  musteriAdres: string;
  musteriSehir: string;
  sube: Sube;
  subeOneri: Sube;
  subeOneriNeden: string;
  planTarih: string | null;   // YYYY-MM-DD — null: planlanmamış kuyruk
  planSaat: string;           // "HH:MM" ya da ""
  planSira: number;           // gün içi sıra
  teslimTarih: string | null; // müşteriye söz verilen gün
  durum: IsDurum;
  kalemler: IsKalem[];
  adet: number;
  tutar: number;
  notlar: string;
  foyYol: string | null;      // üretilmiş/yüklenmiş föy PDF'inin Blob yolu
  foyAt: string | null;
  ekler: IsEk[];
  ham: Record<string, unknown>;  // kaynak ham verisi (ikas siparişi, not metni…)
  olusturan: string;
  createdAt: string;
  updatedAt: string;
  tamamlandiAt: string | null;
  kaynakAt: string | null;    // kaynağın son güncelleme zamanı (senkron kıyası)
}

export type OlayTur = "olustur" | "durum" | "tasi" | "sube" | "not" | "foy" | "ek" | "senk" | "teklif" | "duzenle";

export interface IsOlay {
  id: string;
  isId: string;
  tur: OlayTur;
  aciklama: string;
  by: string;
  at: string;
  meta: Record<string, unknown>;
}

export interface GunNotu {
  id: string;
  tarih: string;        // YYYY-MM-DD
  sube: Sube;           // "" = her iki şube
  metin: string;
  tamam: boolean;
  by: string;
  createdAt: string;
  updatedAt: string;
}

// ---- Teklif ----
export type TeklifDurum = "taslak" | "gonderildi" | "kabul" | "red" | "siparis";

export const TEKLIF_DURUM_LABELS: Record<TeklifDurum, string> = {
  taslak: "Taslak",
  gonderildi: "Gönderildi",
  kabul: "Kabul Edildi",
  red: "Reddedildi",
  siparis: "Siparişe Dönüştü",
};

export interface Teklif {
  id: string;
  no: string;                 // TKL-2026-001
  musteriAd: string;
  musteriTel: string;
  musteriEposta: string;
  musteriAdres: string;
  musteriId: string;          // perakende müşteri defteri kaydı (varsa)
  sube: Sube;
  kalemler: RetailItem[];
  usdKur: number;
  brut: number;
  iskonto: number;
  toplam: number;
  gecerlilik: string | null;  // YYYY-MM-DD
  takipTarih: string | null;  // YYYY-MM-DD — takvimde "ara" hatırlatması
  durum: TeklifDurum;
  notlar: string;
  olusturan: string;
  createdAt: string;
  updatedAt: string;
  siparisRef: string;         // dönüştürülünce PRK numarası
  siparisKey: string;         // dönüştürülünce siparişin dateKey'i
  isId: string;               // dönüştürülünce açılan üretim işi
}

// ---- Takvim yanıtı ----
export interface TakvimYaniti {
  ok: true;
  from: string;
  to: string;
  bugun: string;
  isler: UretimIs[];      // planTarih aralıkta olanlar + planlanmamışlar (plansiz=1)
  notlar: GunNotu[];
  teklifler: Teklif[];    // takipTarih aralıkta olan açık teklifler
  senk?: { perakende?: { eklenen: number; guncellenen: number }; online?: { eklenen: number; guncellenen: number; hata?: string } };
}

// ---- Yardımcılar ----
export const TARIH_RE = /^\d{4}-\d{2}-\d{2}$/;
export const SAAT_RE = /^([01]\d|2[0-3]):[0-5]\d$/;

/** Perakende sipariş durumu ↔ üretim iş durumu eşlemesi. */
export const PERAKENDE_DURUM_ESLE: Record<string, IsDurum> = {
  Beklemede: "planlandi",
  "Hazırlanıyor": "uretimde",
  "Hazır": "hazir",
  "Teslim Edildi": "teslim",
  "İptal": "iptal",
};

export const IS_DURUM_PERAKENDE: Partial<Record<IsDurum, string>> = {
  yeni: "Beklemede",
  planlandi: "Beklemede",
  uretimde: "Hazırlanıyor",
  hazir: "Hazır",
  teslim: "Teslim Edildi",
  tamamlandi: "Teslim Edildi",
  iptal: "İptal",
};

/** Ölçü metni: 500×700 mm → "50×70 cm" */
export function olcuCm(mm?: { w: number; h: number }): string {
  if (!mm || !(mm.w > 0) || !(mm.h > 0)) return "";
  const f = (v: number) => (Math.round(v) / 10).toLocaleString("tr-TR", { maximumFractionDigits: 1 });
  return `${f(mm.w)}×${f(mm.h)} cm`;
}
