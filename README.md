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
| `PATRON_WHATSAPP`, `WHATSAPP_TEMPLATE_SIPARIS`, `WHATSAPP_TEMPLATE_DIL` | — | Her siparişin fiş PDF'i WhatsApp Cloud API ile bu alıcılara **dosya olarak** gider (`05325099442:Özgür Bey,05336610287:Gültekin Bey` — iki nokta sonrası şablondaki hitap); onaylı şablon ad(lar)ı (`siparis_fisi,siparis_fisi_v2`, sırayla denenir) ve dili (`tr`). Kurulum durumu ve test: `/panel/ayarlar` (sahipler) |
| `WHATSAPP_TEMPLATE_MUSTERI`, `WHATSAPP_VERIFY_TOKEN`, `WHATSAPP_APP_SECRET` | — | Müşteriye fiş: sipariş formunda "müşteriye bildir" işaretliyse önce WhatsApp ile PDF (2 değişkenli şablon), Meta "teslim edilemedi" derse webhook (`/api/whatsapp/webhook`) SMS'e düşer. Doğrulama metni ve imza gizi (App Secret; gelen kutusu için zorunlu, yoksa webhook olayları reddedilir) |
| `MIKRO_API_URL`, `MIKRO_API_KEY`, `MIKRO_FIRMA_KODU`, `MIKRO_KULLANICI`, `MIKRO_SIFRE`, `MIKRO_CALISMA_YILI` | — | Mikro Jump 17 Desktop API (yalnızca okuma): cari bakiye sorgusu. Bağlantı denemesi `/panel/ayarlar` (sahipler) |
| `MIKRO_DEPO_ANKARA`, `MIKRO_DEPO_ISTANBUL` | DEPOLAR'daki ada göre | Mikro'dan stok çekerken şubeye sayılacak depo numaraları (`1,3` gibi). Boşsa adında ANKARA / İSTANBUL geçen depolar |
| `STOK_TAZELIK_DK` | 120 | Çalışan stok sorgularken yayındaki veri bu kadar dakikadan eskiyse Mikro'dan tazelenir; ayrıca cron her sabah 07:30'da çeker (`/api/stock/mikro`) |
| `BOLGE_SORUMLULARI` | Ankara: Ramazan Kaypan, İstanbul: Alaattin Yıldız, Taşra: Murat Gündüz | Müşteri satış bölgeleriyle ilgilenen satışçılar: `ankara=…;istanbul=…;tasra=…` |
| `DATABASE_URL` | Mesajlar için | Postgres (Vercel Storage → Neon). `STORAGE_URL` / `POSTGRES_URL` gibi farklı ön ekli adlar da tanınır. Gelen kutusu tabloları ilk açılışta kendiliğinden kurulur |
| `MESAJ_USERNAMES` | bütün çalışanlar | Gelen kutusunu (`/panel/mesajlar`) görüp yanıtlayabilenleri daraltır (virgülle; sahipler her zaman dahil) |
| `WHATSAPP_TEMPLATE_SERBEST` | — | Gelen kutusundan bizim başlattığımız WhatsApp mesajı ve 24 saat sonrası yanıt için Meta onaylı şablon ad(lar)ı (`{{1}}` = metin) |
| `GMAIL_HESAPLAR` | — | `adres:uygulama-şifresi;adres2:şifre2` — bu Gmail hesaplarının gelen kutusu IMAP ile okunur, yanıt aynı hesaptan SMTP ile gider |
| `INSTAGRAM_TOKEN`, `INSTAGRAM_PAGE_ID`, `INSTAGRAM_ACCOUNT_ID` | — | Instagram DM'leri (Meta Messenger Platform); webhook `/api/mesaj/webhook`. `META_APP_SECRET` (ya da `WHATSAPP_APP_SECRET`) zorunlu; isteğe bağlı `INSTAGRAM_VERIFY_TOKEN` |

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

## Mesajlar (Gelen Kutusu)

`/panel/mesajlar` — WhatsApp (Cloud API numarası), Instagram (olga.cerceve) ve
Gmail hesaplarına gelen mesajlar tek listede toplanır; çalışan aynı ekrandan
yanıtlar. Yalnızca `MESAJ_USERNAMES`'teki çalışanlar (ve sahipler) görür.

- **Kayıt:** Postgres (`DATABASE_URL`). Konuşma = kanal + karşı taraf (numara / IGSID / e-posta adresi).
  Telefon veya e-posta müşteri defterindeki (toptan `C…` ya da perakende `P…`) bir kartla
  eşleşirse konuşma o karta bağlanır; eşleşmezse "kayıtsız" kalır, elle bağlanabilir.
- **WhatsApp:** mevcut webhook (`/api/whatsapp/webhook`) gelen mesajları ve teslim/okundu
  durumlarını da yazar. Müşteri yazdığında ve 24 saat içinde yanıtlandığında normal yazışma gibidir.
  **24 saat** geçtiyse ya da konuşmayı **biz başlatıyorsak** (liste başlığındaki "Yeni" düğmesi:
  müşteri defterinden seç ya da numara yaz) Meta onaylı şablon gerekir: `WHATSAPP_TEMPLATE_SERBEST`
  tanımlıysa metin şablonun içinde gider, müşteri yanıtlayınca serbest yazışma açılır. Şablon yoksa
  ekran uyarır ve yalnızca 24 saat içinde yanıt verilir.
  Not: müşterinin başlattığı yazışmalar Meta'da ücretsizdir; bizim başlattığımız şablonlu mesajlar
  Meta'nın mesaj başı ücretine tabidir.
- **Instagram** (olga.cerceve profesyonel hesabı; WhatsApp ile aynı Meta uygulaması, "Instagram" kullanım durumu):
  1. Instagram uygulamasında: Ayarlar → Mesajlar ve hikâye yanıtları → **Bağlı araçlar** → *Mesajlara erişime izin ver* açık.
  2. Meta uygulaması → Add use cases → Business messaging → **Instagram** → Customize → **API setup with Instagram login**
     (önerilen yol): "Add all required permissions"; "Generate access tokens" ile olga.cerceve hesabını ekleyip jetonu ve
     Instagram hesap kimliğini alın; "Configure webhooks": Callback URL `https://<site>/api/mesaj/webhook`, doğrulama metni
     `WHATSAPP_VERIFY_TOKEN`, alan `messages`. Aynı sayfadaki **Instagram app secret** (Show) değerini de alın.
  3. Vercel: `INSTAGRAM_TOKEN`, `INSTAGRAM_ACCOUNT_ID`, `INSTAGRAM_APP_SECRET` → Redeploy. (`INSTAGRAM_PAGE_ID` boş kalır.
     Jeton 60 günlüktür; sistem 7 günde bir kendisi tazeleyip veri tabanında saklar.)
  4. Alternatif "API setup with Facebook login" yolu: sistem kullanıcısı jetonu (süresiz; izinler `instagram_basic`,
     `instagram_manage_messages`, `pages_manage_metadata`, `pages_messaging`, `pages_show_list`) + `INSTAGRAM_PAGE_ID`;
     imza `META_APP_SECRET` / `WHATSAPP_APP_SECRET`; Ayarlar kartındaki "Sayfa aboneliğini onar" gerekir.
  5. Ayarlar → Mesajlar kartı → "Bağlantıları sına": @olga.cerceve ve jeton durumu görünmeli. Kart, Meta
     uygulamasının kendi webhook aboneliğini de (`/{app-id}/subscriptions`, "instagram" nesnesi → `messages`) sorgular;
     eksik ya da adres farklıysa **"Uygulama webhook'unu onar"** düğmesi bizim adresimizle abone yapar (Meta bu sırada
     callback'i `INSTAGRAM_VERIFY_TOKEN` ile doğrular). Sayfa aboneliği tek başına yetmez; ikisi de ✓ olmalı.
  6. Development modunda yalnızca uygulamada rolü olan Instagram hesaplarının (Instagram testers) DM'leri gelir; bütün
     müşteriler için `instagram_business_manage_messages` iznine **App Review** alınıp uygulama **Live** yapılır.
  Instagram uygulamasından atılan yanıtlar da (echo) konuşmada görünür. 24 saat kuralı burada da geçerlidir.
- **Gmail:** hesap başına Google *uygulama şifresi* (2 adımlı doğrulama açık olmalı). Gelen kutusu
  ekran açıkken 60 sn'de bir, ilk kurulumda son 7 gün okunur; `noreply`/bülten adresleri sessiz düşer.
  Yanıt aynı hesaptan, aynı konu dizisine (In-Reply-To) gider. `/api/mesaj/senk` elle/cron ile de tetiklenebilir.
- **Yapay zekâ:** "Taslak öner" düğmesi, konuşmayı + katalog/stok/kur/müşteri kartını okuyup yanıt
  **taslağı** yazar (`ANTHROPIC_API_KEY`). Hiçbir şey kendiliğinden gönderilmez; çalışan okur, düzeltir, gönderir.
  Taslaktan gönderilen mesajlar konuşmada "taslak" etiketiyle görünür.
- **Ekler:** görsel/belge/ses Vercel Blob'a *özel* olarak kaydedilir ve yalnızca oturumlu, yetkili
  kullanıcıya `/api/mesaj/ek` üzerinden gösterilir.

> 0850 305 75 45 numarasındaki *WhatsApp Business uygulaması* bu kutuya bağlı değildir: Meta bir numarayı
> ya telefondaki uygulamada ya da Cloud API'de çalıştırır. O numaranın yazışmalarını da burada görmek için
> numara Cloud API'ye taşınmalıdır (taşınınca telefondaki uygulama o numara için kapanır).

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
