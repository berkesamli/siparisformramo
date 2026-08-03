// Teknik malzeme kataloğu — Form.html'den taşındı.
// priceEUR: Euro fiyatlı ürünler (kutu fiyatı), priceTL: TL fiyatlı ürünler (OLGA).

import type { StockStatus } from "./catalog";

export interface TechnicalProduct {
  code: string;
  name: string;
  category: string;
  adetPerKutu: number;
  priceEUR?: number;
  priceTL?: number;
  isKarton?: boolean;
  stok?: StockStatus;
}

export const TECHNICAL_PRODUCTS: TechnicalProduct[] = [
  // ==================== POZZİ ====================
  { code: "NO1-UC", name: "NO: 1 Üçgen Askı", category: "Pozzi", adetPerKutu: 1000, priceEUR: 17.5 },
  { code: "NO2-UC", name: "NO: 2 Üçgen Askı", category: "Pozzi", adetPerKutu: 1000, priceEUR: 21.0 },
  { code: "NO3-UC", name: "NO: 3 Üçgen Askı", category: "Pozzi", adetPerKutu: 500, priceEUR: 17.5 },
  { code: "CIFT-MENT", name: "Çiftli Menteşe", category: "Pozzi", adetPerKutu: 1000, priceEUR: 55.0 },
  { code: "KUC-YUV", name: "Küçük Yuvarlak Askı", category: "Pozzi", adetPerKutu: 500, priceEUR: 15.5 },
  { code: "BUY-YUV", name: "Büyük Yuvarlak Askı", category: "Pozzi", adetPerKutu: 250, priceEUR: 14.5 },
  { code: "VID-TIRT", name: "Vidalı Tırtıllı Askı", category: "Pozzi", adetPerKutu: 250, priceEUR: 12.0 },
  { code: "BUY-CIV-TIRT", name: "Büyük Çivili Tırtıllı Askı", category: "Pozzi", adetPerKutu: 250, priceEUR: 12.0 },
  { code: "KUC-CIV-TIRT", name: "Küçük Çivili Tırtıllı Askı", category: "Pozzi", adetPerKutu: 500, priceEUR: 20.0 },
  { code: "CIVI-20", name: "20 mm Çivi", category: "Pozzi", adetPerKutu: 100, priceEUR: 4.25 },
  { code: "CIVI-25", name: "25 mm Çivi", category: "Pozzi", adetPerKutu: 100, priceEUR: 4.5 },
  { code: "CIVI-30", name: "30 mm Çivi", category: "Pozzi", adetPerKutu: 100, priceEUR: 4.75 },
  { code: "PULLU-CIVI", name: "Pullu Çivi - Duvar Çivisi", category: "Pozzi", adetPerKutu: 1, priceEUR: 10.0 },
  { code: "TEK-DEL-BUY", name: "Tek Delikli Büyük Askı", category: "Pozzi", adetPerKutu: 100, priceEUR: 15.0 },
  { code: "CIFT-DEL-BUY", name: "Çift Delikli Büyük Askı", category: "Pozzi", adetPerKutu: 100, priceEUR: 19.5 },
  { code: "UC-DEL-BUY", name: "Üç Delikli Büyük Askı", category: "Pozzi", adetPerKutu: 50, priceEUR: 16.5 },
  { code: "ARKALIK-KLIPS", name: "Arkalık Klipsi", category: "Pozzi", adetPerKutu: 100, priceEUR: 5.25 },
  { code: "FOTOBLOK-ASKI", name: "Fotoblok Askısı", category: "Pozzi", adetPerKutu: 500, priceEUR: 50.0 },
  { code: "BUY-CAM-KLIPS", name: "Büyük Cam Klipsi", category: "Pozzi", adetPerKutu: 500, priceEUR: 55.0 },
  { code: "KUC-CAM-KLIPS", name: "Küçük Cam Klipsi", category: "Pozzi", adetPerKutu: 1000, priceEUR: 65.0 },
  // ==================== ALFAMACCHINE ====================
  { code: "ALFA-5", name: "Alfa 5'lik Agraf", category: "Alfamacchine", adetPerKutu: 5000, priceEUR: 16.5 },
  { code: "ALFA-7", name: "Alfa 7'lik Agraf", category: "Alfamacchine", adetPerKutu: 4000, priceEUR: 13.5 },
  { code: "ALFA-10", name: "Alfa 10'luk Agraf", category: "Alfamacchine", adetPerKutu: 3000, priceEUR: 10.5 },
  { code: "ALFA-12", name: "Alfa 12'lik Agraf", category: "Alfamacchine", adetPerKutu: 3000, priceEUR: 12.5 },
  { code: "ALFA-15", name: "Alfa 15'lik Agraf", category: "Alfamacchine", adetPerKutu: 2000, priceEUR: 10.0 },
  { code: "ESNER-CIVI", name: "Esner Çivi", category: "Alfamacchine", adetPerKutu: 12000, priceEUR: 33.0 },
  { code: "YEDEK-PEDAL", name: "Yedek Pedal", category: "Alfamacchine", adetPerKutu: 1, priceEUR: 250.0 },
  { code: "MINIGRAF-U200", name: "Minigraf U200", category: "Alfamacchine", adetPerKutu: 1, priceEUR: 1500.0 },
  { code: "ALFA-U200P", name: "Alfa U-200P", category: "Alfamacchine", adetPerKutu: 1, priceEUR: 2000.0 },
  { code: "ALFA-U300P", name: "Alfa U300P", category: "Alfamacchine", adetPerKutu: 1, priceEUR: 2400.0 },
  { code: "T-200", name: "T-200", category: "Alfamacchine", adetPerKutu: 1, priceEUR: 16500.0 },
  { code: "T-400", name: "T-400", category: "Alfamacchine", adetPerKutu: 1, priceEUR: 16500.0 },
  // ==================== CASSESE ====================
  { code: "CASS-5", name: "Cassese 5'lik Agraf", category: "Cassese", adetPerKutu: 8000, priceEUR: 35.0 },
  { code: "CASS-7", name: "Cassese 7'lik Agraf", category: "Cassese", adetPerKutu: 8000, priceEUR: 35.0 },
  { code: "CASS-10", name: "Cassese 10'luk Agraf", category: "Cassese", adetPerKutu: 8000, priceEUR: 35.0 },
  { code: "CASS-12", name: "Cassese 12'lik Agraf", category: "Cassese", adetPerKutu: 8000, priceEUR: 45.0 },
  { code: "CASS-15", name: "Cassese 15'lik Agraf", category: "Cassese", adetPerKutu: 8000, priceEUR: 60.0 },
  { code: "CASS-F15-ESNER", name: "Cassese F-15 Esner Çivi", category: "Cassese", adetPerKutu: 15000, priceEUR: 45.0 },
  { code: "CASS-F15-TAB", name: "Cassese F-15 Tabanca", category: "Cassese", adetPerKutu: 1, priceEUR: 100.0 },
  { code: "CASS-5-KART", name: "Cassese 5'lik Agraf Kartuşlu", category: "Cassese", adetPerKutu: 6, priceEUR: 24.0 },
  { code: "CASS-7-KART", name: "Cassese 7'lik Agraf Kartuşlu", category: "Cassese", adetPerKutu: 6, priceEUR: 26 },
  { code: "CASS-10-KART", name: "Cassese 10'luk Agraf Kartuşlu", category: "Cassese", adetPerKutu: 6, priceEUR: 28.0 },
  { code: "CASS-12-KART", name: "Cassese 12'lik Agraf Kartuşlu", category: "Cassese", adetPerKutu: 6, priceEUR: 30.0 },
  { code: "CASS-15-KART", name: "Cassese 15'lik Agraf Kartuşlu", category: "Cassese", adetPerKutu: 6, priceEUR: 32.0 },
  // ==================== DANLIST ====================
  { code: "MORSO-GIYOTIN", name: "Morso-F Giyotin", category: "Danlist", adetPerKutu: 1, priceEUR: 4000.0 },
  { code: "MORSO-BICAK", name: "Morso Yedek Bıçak", category: "Danlist", adetPerKutu: 1, priceEUR: 350.0 },
  // ==================== RO-MA MAESTRİ ====================
  { code: "F18-MEK", name: "F-18 Çivi Tabancası Mekanik", category: "Ro-ma Maestri", adetPerKutu: 1, priceEUR: 110.0 },
  { code: "F18-HAV", name: "F-18 Çivi Tabancası Havalı", category: "Ro-ma Maestri", adetPerKutu: 1, priceEUR: 200.0 },
  { code: "F15-MEK", name: "F-15 Çivi Tabancası Mekanik", category: "Ro-ma Maestri", adetPerKutu: 1, priceEUR: 110.0 },
  { code: "F15-HAV", name: "F-15 Çivi Tabancası Havalı", category: "Ro-ma Maestri", adetPerKutu: 1, priceEUR: 200.0 },
  { code: "F15-ELEK", name: "Elektrikli F-15 Çivi Tabancası", category: "Ro-ma Maestri", adetPerKutu: 1, priceEUR: 200.0 },
  { code: "DUOFIX-P53", name: "DUOFIX P53", category: "Ro-ma Maestri", adetPerKutu: 1, priceEUR: 30.0 },
  { code: "ELET-F18-5K", name: "Elet.F/-18 Sert Çivi 5000", category: "Ro-ma Maestri", adetPerKutu: 5000, priceEUR: 11.0 },
  { code: "ELET-F18-10K", name: "Elet.F/-18 Sert Çivi 10000", category: "Ro-ma Maestri", adetPerKutu: 10000, priceEUR: 22.0 },
  // ==================== SCAPPİ CARTONİ ====================
  { code: "SCAPPI-DUZ", name: "Scappi Düz Renkler", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 5.5, isKarton: true },
  { code: "SCAPPI-DAMARLI", name: "Scappi Damarlı Renkler", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 8.5, isKarton: true },
  { code: "SCAPPI-ALTIN", name: "Scappi Altın-Gümüş", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 13.5, isKarton: true },
  { code: "SCAPPI-KADIFE", name: "Scappi Kadife", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 20.0, isKarton: true },
  { code: "SCAPPI-TELALI", name: "Scappi Telalı", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 20.0, isKarton: true },
  { code: "SCAPPI-DOKULU", name: "Scappi Dokulu", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 14.0, isKarton: true },
  { code: "SCAPPI-YALDIZ", name: "Scappi Yaldız Dokulu", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 15.0, isKarton: true },
  { code: "SCAPPI-DAM-KART", name: "Scappi Damarlı-Dokulu Karton", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 15, isKarton: true },
  { code: "SCAPPI-JUMBO", name: "Scappi Jumbo Karton", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 18.0, isKarton: true },
  { code: "FOTOBLOK-5MM", name: "Fotoblok 5mm (101x152cm)", category: "Scappi Cartoni", adetPerKutu: 1, priceEUR: 18.0 },
  // ==================== OLGA (TL FİYATLI) ====================
  { code: "ASKI-TELI", name: "Askı Teli", category: "OLGA", adetPerKutu: 1, priceTL: 1000 },
  { code: "YERLI-TIRT", name: "Yerli Tırtıllı Askı", category: "OLGA", adetPerKutu: 500, priceTL: 650 },
  { code: "BUY-YERLI-DEP", name: "Büyük Yerli Deprem Askı", category: "OLGA", adetPerKutu: 500, priceTL: 900 },
  { code: "YUKSUK", name: "Yüksük", category: "OLGA", adetPerKutu: 1000, priceTL: 750 },
  { code: "STRECH", name: "Strech", category: "OLGA", adetPerKutu: 1, priceTL: 70 },
  { code: "STRECH-TOP", name: "Strech Topağı", category: "OLGA", adetPerKutu: 1, priceTL: 100 },
  { code: "OLUKLU-MUK", name: "Oluklu Mukavva", category: "OLGA", adetPerKutu: 1, priceTL: 160 },
  { code: "YUZUKSUZ-ASKI", name: "Yüzüksüz Askı", category: "OLGA", adetPerKutu: 500, priceTL: 500 },
  { code: "KUC-YERLI-DEP", name: "Küçük Yerli Deprem", category: "OLGA", adetPerKutu: 500, priceTL: 800 },
  { code: "ITHAL-5-AGRAF", name: "İthal 5'lik Agraf", category: "OLGA", adetPerKutu: 12000, priceTL: 1900 },
  { code: "ITHAL-7-AGRAF", name: "İthal 7'lik Agraf", category: "OLGA", adetPerKutu: 6000, priceTL: 950 },
  { code: "ITHAL-10-AGRAF", name: "İthal 10'luk Agraf", category: "OLGA", adetPerKutu: 6000, priceTL: 950 },
  { code: "ITHAL-12-AGRAF", name: "İthal 12'lik Agraf", category: "OLGA", adetPerKutu: 6000, priceTL: 950 },
  { code: "ITHAL-15-AGRAF", name: "İthal 15'lik Agraf", category: "OLGA", adetPerKutu: 6000, priceTL: 950 },
  { code: "ITHAL-F15-TAB", name: "İthal F-15 Tabanca", category: "OLGA", adetPerKutu: 1, priceTL: 6500 },
  { code: "ITHAL-F15-ESNER", name: "İthal F-15 Esner Çivi", category: "OLGA", adetPerKutu: 10000, priceTL: 950 },
  { code: "ZIMBA-TELI", name: "Zımba Teli", category: "OLGA", adetPerKutu: 1, priceTL: 100 },
  { code: "GAZLI-ELMA", name: "Gazlı Elmas", category: "OLGA", adetPerKutu: 1, priceTL: 600 },
  { code: "ITHAL-NO1-ASKI", name: "İthal No:1 Askı", category: "OLGA", adetPerKutu: 1000, priceTL: 750 },
  { code: "KUC-CIV-TIRT-OLGA", name: "Küçük Çivili Tırtıllı Askı", category: "OLGA", adetPerKutu: 1000, priceTL: 1300 },
  { code: "TIRTILLI-ASKI", name: "Tırtıllı Askı", category: "OLGA", adetPerKutu: 500, priceTL: 650 },
  { code: "BUY-TIRT-VID", name: "Büyük Tırtıllı Askı Vidalı", category: "OLGA", adetPerKutu: 300, priceTL: 600 },
  { code: "KUC-TIRT-VID", name: "Küçük Tırtıllı Askı Vidalı", category: "OLGA", adetPerKutu: 500, priceTL: 750 },
  { code: "ITHAL-TEK-DEL-BUY", name: "İthal Tek Delikli Büyük Askı", category: "OLGA", adetPerKutu: 100, priceTL: 550 },
  { code: "ITHAL-CIFT-DEL-BUY", name: "İthal Çift Delikli Büyük Askı", category: "OLGA", adetPerKutu: 100, priceTL: 650 },
  { code: "ITHAL-UC-DEL-BUY", name: "İthal Üç Delikli Büyük Askı", category: "OLGA", adetPerKutu: 50, priceTL: 550 },
  { code: "ITHAL-TEK-DEL-KUC", name: "İthal Tek Delikli Küçük Askı", category: "OLGA", adetPerKutu: 100, priceTL: 300 },
  { code: "ITHAL-CIFT-DEL-KUC", name: "İthal Çift Delikli Küçük Askı", category: "OLGA", adetPerKutu: 100, priceTL: 350 },
  { code: "KUC-VIDA-1KG", name: "Küçük Vida (1kg)", category: "OLGA", adetPerKutu: 1, priceTL: 350 },
  { code: "BUY-VIDA-1KG", name: "Büyük Vida (1kg)", category: "OLGA", adetPerKutu: 1, priceTL: 350 },
  { code: "KAGIT-BANT", name: "Kağıt Bant (3,6cm)", category: "OLGA", adetPerKutu: 1, priceTL: 125 },
  { code: "CIFT-TAR-BANT", name: "Çift Taraflı Bant (1,5cm)", category: "OLGA", adetPerKutu: 1, priceTL: 60 },
  { code: "KOLI-BANTI", name: "Koli Bantı", category: "OLGA", adetPerKutu: 1, priceTL: 150 },
  { code: "SIYAH-KRAFT", name: "Siyah Kraft Bant (6cm)", category: "OLGA", adetPerKutu: 1, priceTL: 125 },
  { code: "TUVAL-PENS", name: "Tuval Pensesi", category: "OLGA", adetPerKutu: 1, priceTL: 650 },
  { code: "KUC-DUVAR-ASKI", name: "Küçük Duvar Askı Aparatı", category: "OLGA", adetPerKutu: 100, priceTL: 550 },
  { code: "ORTA-DUVAR-ASKI", name: "Orta Duvar Askı Aparatı", category: "OLGA", adetPerKutu: 50, priceTL: 400 },
  { code: "BUY-DUVAR-ASKI", name: "Büyük Duvar Askı Aparatı", category: "OLGA", adetPerKutu: 50, priceTL: 450 },
  { code: "BUY-MAKET", name: "Büyük Maket Bıçağı", category: "OLGA", adetPerKutu: 1, priceTL: 90 },
  { code: "BUY-MAKET-UCU", name: "Büyük Maket Bıçağı Ucu", category: "OLGA", adetPerKutu: 1, priceTL: 150 },
  { code: "KUC-MAKET", name: "Küçük Maket Bıçağı", category: "OLGA", adetPerKutu: 1, priceTL: 75 },
  { code: "KUC-MAKET-UCU", name: "Küçük Maket Bıçağı Ucu", category: "OLGA", adetPerKutu: 1, priceTL: 125 },
  { code: "ITHAL-FOTOBLOK-ASKI", name: "İthal Fotoblok Askısı", category: "OLGA", adetPerKutu: 100, priceTL: 600 },
  { code: "KOMPRESOR-KON", name: "Kompresör Konnektörü", category: "OLGA", adetPerKutu: 1, priceTL: 200 },
  { code: "MDF-ARKALIK", name: "MDF Arkalık Askısı", category: "OLGA", adetPerKutu: 100, priceTL: 100 },
  { code: "KOSE-5X3", name: "Plastik Köşe Koruyucu (5x3cm)", category: "OLGA", adetPerKutu: 1, priceTL: 7.5 },
  { code: "KOSE-6X4", name: "Plastik Köşe Koruyucu (6x4cm)", category: "OLGA", adetPerKutu: 1, priceTL: 8.0 },
  { code: "KOSE-6X5", name: "Plastik Köşe Koruyucu (6x5cm)", category: "OLGA", adetPerKutu: 1, priceTL: 9.5 },
  // ==================== NS SERİSİ ====================
  { code: "NS-KARTON-DUZ", name: "NS Karton Düz", category: "NS Serisi", adetPerKutu: 1, priceTL: 200, isKarton: true },
  { code: "NS-KARTON-ALTIN", name: "NS Karton Altın-Gümüş", category: "NS Serisi", adetPerKutu: 1, priceTL: 350, isKarton: true },
  { code: "NS-KARTON-KADIFE", name: "NS Karton Kadife", category: "NS Serisi", adetPerKutu: 1, priceTL: 600, isKarton: true },
];

export function getTechnicalProduct(code: string): TechnicalProduct | undefined {
  return TECHNICAL_PRODUCTS.find((t) => t.code === code);
}

export function technicalByCategory(): Record<string, TechnicalProduct[]> {
  const map: Record<string, TechnicalProduct[]> = {};
  for (const t of TECHNICAL_PRODUCTS) {
    (map[t.category] ||= []).push(t);
  }
  return map;
}
