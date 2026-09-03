# Olga Çerçeve — Bayi Paneli

Olga Çerçeve **toptan bayileri** için ayrı bir web uygulaması. Bayi, kendi
perakende müşterisine Olga'nın online çerçeve hesaplayıcısıyla anında fiyat
verir, siparişi kaydeder, müşteriye PDF / WhatsApp gönderir ve siparişlerini
takip eder. Amaç: bayinin işini kolaylaştırarak Olga profil satışını artırmak.

> Bu uygulama `olgasiparis.com` (sipariş/finans platformu) ile **ayrı** çalışır.
> Ortak nokta yalnızca çerçeve kataloğu ve fiyat hesabı mantığıdır.

## Kim ne yapar

| Rol | Giriş | Yapabildikleri |
|---|---|---|
| **Olga yöneticisi** | `ADMIN_USERNAME` / `ADMIN_PASSWORD` | Bayi açar/kapatır, abonelik durumunu ve muafiyeti belirler, kullanım istatistiklerini görür (`/yonetim`) |
| **Bayi** | Yöneticinin verdiği kullanıcı adı | Fiyat çarpanlarını ayarlar, online çerçeve sihirbazıyla teklif/sipariş oluşturur, siparişleri takip eder (`/panel`) |
| **Bayinin müşterisi** | Giriş yok | WhatsApp'tan gelen imzalı takip linkiyle siparişinin durumunu ve fişini görür (`/takip`) |

## Fiyatlandırma modeli

- **Çerçeve:** Olga toptan liste fiyatı (USD/mt, `data/catalog.ts`) × **bayi çarpanı** × USD kuru (TCMB otomatik veya sabit)
- **Paspartu / Cam:** bayi başına TL/m² · **Baskı:** bayi başına USD/m² · **İşçilik:** kalem başına sabit TL (opsiyonel)
- Hesap kuralı Olga perakende ile aynıdır (`data/pricing.ts` → `computeCosts`).
- Yeni bayi Olga perakende varsayılanlarıyla başlar (çarpan 5, cam 2.000 ₺/m² vb.) ve panelden değiştirir.
- Sihirbazda bayi, çerçeve kaleminin toptan liste maliyetini ve brüt farkını görür (müşteriye gösterilmez).

## Abonelik

Bayi kaydında `subscription.status` dört değer alır:

| Durum | Anlamı | Etkisi |
|---|---|---|
| `aktif` | Aylık ücret ödüyor | Tam kullanım |
| `muaf` | Aylık toptan alımı eşiği geçti → ücretsiz | Tam kullanım |
| `odeme_bekliyor` | Ödeme gecikti | `paidUntil` + 7 gün sonra yeni sipariş kaydı kapanır |
| `askida` | Askıya alındı | Giriş yapar, fiyat hesaplar, ama sipariş kaydedemez |
| hesap `active=false` | Pasif | Giriş yapamaz |

Kart tahsilatı (iyzico / PayTR) bu sürümde yoktur; yönetici ödemeyi elle işaretler.
Muafiyet, ana sitedeki toptan alım verisiyle ileride otomatik bağlanacaktır.

## Kurulum (Vercel)

1. Bu klasörü **ayrı bir GitHub reposuna** taşıyın ve Vercel'de yeni proje olarak açın.
2. Vercel → Storage → **Blob** oluşturup projeye bağlayın (`BLOB_READ_WRITE_TOKEN` otomatik gelir).
   Blob olmadan bayi/sipariş/ayar kaydı yapılamaz.
3. `.env.example` içindeki değişkenleri girin: en az `AUTH_SECRET`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`.
4. Deploy edin; `/giris` → yönetici hesabıyla girip `/yonetim` sayfasından ilk bayiyi tanımlayın.
5. Alan adı: örn. `bayi.olgacerceve.com`.

## Yerel geliştirme

```bash
npm install
cp .env.example .env.local   # değerleri doldurun
npm run dev                  # http://localhost:3000
```

## Veri düzeni (Vercel Blob)

```
bayiler/_index.json                       bayi listesi (giriş için)
bayiler/<slug>.json                       bayi kaydı (şifre scrypt özeti, abonelik)
bayiler/<slug>/fiyat.json                 bayinin fiyat ayarları
bayiler/<slug>/siparisler/<tarih>/<no>.json  siparişler
bayiler/<slug>/sayac-<yıl>.json           sipariş numarası sayacı
kur/<tarih>.json                          TCMB USD kuru (günlük önbellek)
```

Depolama katmanı `lib/store.ts` içinde tek noktadadır; bayi sayısı büyüyünce
Postgres'e (Vercel Postgres / Supabase) geçiş yalnızca bu dosyayı ve
`lib/dealers.ts` / `lib/orders.ts` yol fonksiyonlarını etkiler.

## Katalog güncelleme

Çerçeve profilleri ve toptan liste fiyatları `data/catalog.ts`, profil
görselleri `data/frame-images.ts` dosyasındadır (ana siteyle aynı içerik).
Ana site güncellendiğinde bu dosyaları kopyalayıp push edin. Sonraki adım:
ana siteye token korumalı bir `/api/bayi/katalog` ucu ekleyip buradan günlük
çekmek.

## Yol haritası

- [ ] iyzico / PayTR ile aylık kart tahsilatı ve otomatik `odeme_bekliyor` geçişi
- [ ] Ana siteden bayi bazında aylık toptan alım cirosu senkronu → otomatik muafiyet
- [ ] Ana siteden katalog / stok senkronu
- [ ] Bayiye özel herkese açık tasarım sayfası (müşteri kendi evinden çerçeve seçer, talep WhatsApp'a düşer)
- [ ] Bayi logosu (PDF ve fişte)
- [ ] Bayi altında birden fazla kullanıcı
