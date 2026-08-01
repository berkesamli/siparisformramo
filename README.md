# Olga Çerçeve — Sipariş ve Katalog Platformu

Google Apps Script tabanlı sipariş formunun modern, Vercel'de yayınlanabilir
**Next.js** sürümü. Tek sitede dört işlev:

| Bölüm | Kim kullanır | Ne yapar |
|---|---|---|
| `/panel` | Çalışanlar | Sipariş oluşturur; sipariş **e-posta + WhatsApp** ile firmaya iletilir |
| `/portal` | Müşteriler/Bayiler | Ürünleri, stok durumunu ve **toptan fiyat listesini** görür |
| `/kataloglar` | Herkes | PDF katalogları **dergi görünümünde** (sayfa çevirmeli) inceler |
| 🤖 Asistan | Giriş yapanlar | Claude API destekli ürün/fiyat asistanı |

Eski Apps Script kodu `legacy-apps-script/` klasöründe korunmaktadır.

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
