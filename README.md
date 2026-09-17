# Olga Çerçeve — Sipariş ve Katalog Platformu

Google Apps Script tabanlı sipariş formunun modern, Vercel'de yayınlanabilir
**Next.js** sürümü. Koyu "cam" temalı, sol menülü bir yönetim paneli olarak
çalışır; telefon, tablet ve masaüstü için ayrı ayrı düzenlenmiştir.

| Bölüm | Kim kullanır | Ne yapar |
|---|---|---|
| `/` | Giriş yapanlar | **Gösterge paneli**: bugünkü/açık siparişler, 14 günlük akış grafiği, durum dağılımı, son siparişler, uyarılar, hızlı işlemler |
| `/panel` | Çalışanlar | Sipariş oluşturur; sipariş **e-posta + WhatsApp** ile firmaya iletilir |
| `/portal` | Müşteriler/Bayiler | Ürünleri, stok durumunu ve **toptan fiyat listesini** görür |
| `/kataloglar` | Herkes | PDF katalogları **dergi görünümünde** (sayfa çevirmeli) inceler |
| 🤖 Asistan | Giriş yapanlar | Claude API destekli ürün/fiyat asistanı |

Eski Apps Script kodu `legacy-apps-script/` klasöründe korunmaktadır.

## Arayüz

- **Kabuk** (`components/shell/`): sol kenar çubuğu (masaüstünde daraltılabilir ray,
  tablette ikon rayı + çekmece, telefonda çekmece + alt sekme çubuğu), üst çubuk
  (sayfa yolu, arama, bildirim zili, tema, kullanıcı menüsü). Menü/kırıntı/arama
  tek kaynaktan gelir: `components/shell/nav-config.ts`.
- **Genel arama** `Ctrl/⌘ + K`: sayfalar, toptan ve perakende siparişler, müşteriler,
  stok kodları, çerçeve profilleri ve teknik malzeme (`/api/search`). Müşteriler
  yalnızca ürün/stok arar.
- **Tema**: varsayılan açık tema; üst çubuktan koyu "cam" temaya geçilir, tercih
  tarayıcıda saklanır. Renkler yalnızca `app/styles/tokens.css` değişkenlerinden gelir.
- **Bildirim zili**: yeni sipariş düştüğünde ses + tarayıcı bildirimi (90 sn'de bir
  iki küçük sayaç dosyası okunur).
- **Gösterge paneli** verisi `/api/dashboard` (aylık indeksler + tek stok dosyası +
  günün kuru; 45 sn süreç içi önbellek). Kenar çubuğu sayaçları `?lite=1` ile gelir.
- Stil aileleri `app/styles/` altında: `tokens` → `base` → `shell` → `wizard`,
  `labels`, `pickers`, `orders`, `reports`, `modals`, `dashboard`.
- **Satışlarım** (`/panel/satislarim`, `/api/cirom`): her çalışan yalnızca kendi adına
  girilen siparişlerin cirosunu görür (toptan net + perakende toplam, iptaller hariç);
  bugün / son 7 gün / aylık, 6 aylık grafik, 14 günlük akış ve kendi sipariş listesi.
  Başka çalışanın rakamı hiçbir uçtan dönmez; ad oturumdan alınır.
- **Duyurular** (`data/duyurular.ts`): tarih aralığı ve role göre yayınlanan site içi
  duyurular; ilk girişte bir kez pencere olarak çıkar, sonrasında zil menüsünde kalır.
  Yeni duyuru eklemek için listeye benzersiz `id` ile kayıt eklemek yeterlidir.
- Telefonda "Ana ekrana ekle" ile uygulama gibi açılır (`app/manifest.ts`).

## Vercel'e Yayınlama

1. Bu repoyu GitHub'a push edin (zaten GitHub'da).
2. [vercel.com](https://vercel.com) → **Add New Project** → bu repoyu seçin.
   Framework otomatik olarak Next.js algılanır, ayar gerekmez.
3. **Environment Variables** bölümüne `.env.example` dosyasındaki değişkenleri
   girin (en azından `AUTH_SECRET`).
4. **Deploy** butonuna basın. Siteniz `xxx.vercel.app` adresinde yayında olur;
   isterseniz kendi alan adınızı (örn. `siparis.olgacerceve.com`) bağlayın.

## Ortam Değişkenleri

| Değişken | Zorunlu | Açıklama |
|---|---|---|
| `AUTH_SECRET` | ✅ | Oturum çerezlerini imzalayan gizli anahtar |
| `USERS_JSON` | — | Kullanıcı listesi (girilmezse `data/users.ts` içindeki varsayılanlar kullanılır — **üretimde mutlaka değiştirin**) |
| `SMTP_HOST/PORT/USER/PASS/FROM` | — | Sipariş e-postası için SMTP (Gmail: uygulama şifresi) |
| `ORDER_EMAIL_TO` | — | Sipariş e-postasının gideceği adres |
| `WHATSAPP_TOKEN`, `WHATSAPP_PHONE_ID`, `WHATSAPP_TO` | — | Meta WhatsApp Cloud API — tanımlıysa sipariş otomatik WhatsApp'a düşer; tanımlı değilse panelde tek tıkla **wa.me** linki üretilir |
| `ANTHROPIC_API_KEY` | — | AI ürün asistanı için Claude API anahtarı |
| `PATRON_WHATSAPP`, `WHATSAPP_TEMPLATE_SIPARIS`, `WHATSAPP_TEMPLATE_DIL` | — | Her siparişin fiş PDF'i WhatsApp Cloud API ile bu numara(lar)a **dosya olarak** gider; onaylı şablon ad(lar)ı (`siparis_fisi,siparis_fisi_v2`, sırayla denenir) ve dili (`tr`). Kurulum durumu ve test: `/panel/ayarlar` (sahipler) |

## Kullanıcılar ve Roller

- **staff** (çalışan): sipariş paneli + portal + kataloglar
- **customer** (müşteri/bayi): portal (stok + toptan fiyat listesi) + kataloglar

Varsayılan demo hesaplar `data/users.ts` içindedir (`ramazan/olga2025`,
`musteri/olga123` vb.). Üretimde `USERS_JSON` ortam değişkeni ile gerçek
hesaplarınızı tanımlayın.

## PDF Katalog Ekleme

PDF dosyalarınızı `public/catalogs/` klasörüne koyup push edin:

```
public/catalogs/toptan-fiyat-listesi.pdf
public/catalogs/teknik-malzeme-katalogu.pdf
```

Her PDF, `/kataloglar` sayfasında otomatik listelenir ve dergi görünümünde
(çift sayfa, çevirme animasyonlu) açılır.

## Sipariş Akışı

1. Çalışan panelde satırları girer (çerçeve profili / cam / ayna / teknik
   malzeme / diğer). Fiyat hesaplamaları eski formdaki mantığın birebir
   portudur (metre/boy/koli çevrimi, kur, iskonto, KDV %20).
2. "Siparişi Gönder" → `/api/orders`:
   - SMTP tanımlıysa **e-posta** gönderilir (tablo + toplamlar).
   - WhatsApp Cloud API tanımlıysa **WhatsApp mesajı** otomatik gider;
     değilse panelde hazır metinli **wa.me linki** çıkar.

## Günlük Stok Güncelleme (Excel)

Çalışanlar `/panel/stok` sayfasından muhasebe programının günlük stok
Excel'ini (xls/xlsx) yükler:

- Sistem `DEPO ADI / STOK İSMİ / MİKTAR` kolonlarını otomatik bulur.
- Şimdilik yalnızca çerçeve profilleri ("PROFİL" içeren satırlar) alınır.
- Ankara ve İstanbul depoları ayrı gösterilir; metraj **2,9'a bölünerek boy**
  cinsinden yayınlanır.
- Portaldaki "Güncel Stok Sorgula" sekmesi bulanık arama yapar:
  `gc065-1473` yazan biri `GC065-1473BX`i bulur; küçük yazım hataları da
  tolere edilir.

Kalıcı saklama için Vercel'de bir kez **Storage → Create Database → Blob**
oluşturup projeye bağlayın (`BLOB_READ_WRITE_TOKEN` otomatik tanımlanır).
Blob yapılandırılana kadar site, repo içindeki `data/stock-snapshot.json`
dosyasını gösterir.

## Ürün / Fiyat Güncelleme

- Çerçeve profilleri ve stok durumu: `data/catalog.ts` (`stok: "var" | "az" | "yok"`)
- Teknik malzemeler: `data/technical.ts`
- Cam/ayna plaka ölçüleri: `data/glass.ts`

Değişikliği push ettiğinizde Vercel otomatik yeniden yayınlar.

## Yerel Geliştirme

```bash
npm install
npm run dev   # http://localhost:3000
```

<!-- deploy: 2026-08-01T08:25:44Z -->
