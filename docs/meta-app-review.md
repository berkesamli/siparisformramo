# Meta App Review — Instagram mesajlaşma (siparisformramo, uygulama 1806991880297661)

Meta kuralı: **Standard Access** ile yalnızca uygulamada rolü olan hesapların (yönetici / geliştirici / test
kullanıcısı) DM'leri gelir. Bütün müşterilerin mesajları için `instagram_manage_messages` izninde **Advanced
Access** gerekir; bu da **İşletme Doğrulaması** + **Uygulama İncelemesi** demektir.

## 1. İşletme doğrulaması (Business Verification)

Business Suite → İşletme ayarları → **Güvenlik Merkezi** → "Doğrulamayı başlat".
İstenenler: yasal işletme adı (Olga Çerçeve Sanayi ve Ticaret Ltd. Şti.), adres, telefon, web sitesi
(olgasiparis.com ya da olgacerceve.com) ve bir belge: vergi levhası, ticaret sicil gazetesi ya da imza sirküleri.
Adres ve telefonun belgedekiyle aynı olması gerekir. Süre: genelde 1–3 iş günü.

## 2. Uygulama incelemesi (App Review)

Meta uygulaması → **App Review → Permissions and Features** (ya da "Manage messaging & content on Instagram"
kullanım durumu → izinler). Şu izinlerde **Request advanced access**:

| İzin | Neden |
|---|---|
| `instagram_manage_messages` | DM'leri almak ve yanıtlamak (asıl izin) |
| `instagram_basic` | Gönderenin adı / kullanıcı adı |
| `pages_manage_metadata` | Webhook aboneliği (sayfa) |
| `pages_messaging` | Messenger yolu üzerinden mesajlaşma altyapısı |

Ardından **Submit for review**. Her izin için üç şey istenir: kullanım açıklaması, adım adım anlatım ve
ekran kaydı. Aşağıdaki İngilizce metinler kopyalanabilir.

### Uygulama ayarları (önce tamamlayın)

- App settings → Basic: **Privacy Policy URL** `https://olgasiparis.com/gizlilik`,
  **User data deletion** → "Data deletion instructions URL" `https://olgasiparis.com/gizlilik#data-deletion`,
  App icon (1024×1024, Olga logosu), Category: *Business and pages*, Contact email.
- App Mode / Publish: yayınlanmış olmalı (yapıldı).

### Ortak açıklama (her izne aynı paragraf)

> This app is an internal order-management tool used only by Olga Çerçeve (a picture-frame manufacturer in
> Ankara, Türkiye) for its own business assets: the Facebook Page "Olga Çerçeve" and the Instagram professional
> account @olga.cerceve, both owned by our verified business. It is not offered to any third party. Staff members
> log in to our website (olgasiparis.com) and see customer messages from Instagram, WhatsApp and e-mail in one
> inbox, then reply to the customer from there. Access tokens are generated for a System User in our own Business
> Manager; there is no Facebook Login flow for end users.

### `instagram_manage_messages`

> **How we use it:** We subscribe to the `messages` webhook of our own Instagram professional account
> (@olga.cerceve). When a customer sends us a DM, the message text and attachments are stored in our inbox and shown
> to our staff. Staff reply from the inbox; the reply is sent with `POST /{page-id}/messages` within the 24-hour
> window (messaging_type RESPONSE). We also read `/{page-id}/conversations?platform=instagram` to make sure no
> message is missed. Messages are used only to answer the customer and prepare their order.
>
> **Step-by-step:** 1) A customer sends a DM to @olga.cerceve from the Instagram app. 2) Meta delivers the webhook
> to https://olgasiparis.com/api/mesaj/webhook. 3) The message appears in our inbox at olgasiparis.com/panel/mesajlar
> (staff login required). 4) A staff member types a reply and clicks "Gönder". 5) The customer receives the reply in
> Instagram. The screencast shows exactly these steps.

### `instagram_basic`

> **How we use it:** To show the customer's name and username next to their messages in our inbox, we read the
> basic profile (name, username) of the Instagram user who messaged us. Nothing else is read or published.

### `pages_manage_metadata`

> **How we use it:** To subscribe our own Facebook Page to the `messages` webhook field
> (`POST /{page-id}/subscribed_apps`) so that Instagram messages of the connected professional account are
> delivered to our server. No Page content is modified.

### `pages_messaging`

> **How we use it:** Instagram messaging for a professional account connected to a Facebook Page runs over the
> Messenger Platform; this permission is required by Meta for sending replies through `/{page-id}/messages`.
> We only reply to customers who messaged us first, within the 24-hour window.

### Ekran kaydı senaryosu (1–2 dakika, ses gerekmez)

1. Telefonda başka bir Instagram hesabından (bir çalışanın hesabı) @olga.cerceve'ye DM atın: "Merhaba, 50x70
   çerçeve fiyatı nedir?" — ekranı kaydedin ya da bilgisayara yansıtın.
2. Bilgisayarda olgasiparis.com → giriş → **Mesajlar**. Mesajın listede belirdiğini, konuşmaya tıklayınca metnin
   ve gönderen adının göründüğünü gösterin.
3. Yanıt kutusuna yazıp **Gönder**'e basın (ör. "Merhaba, 50x70 için 1.450 ₺. Mağazamıza bekleriz.").
4. Telefonda yanıtın Instagram'da geldiğini gösterin.
5. İsteğe bağlı: Ayarlar → Mesajlar kartında Instagram satırını (abonelik ✓) gösterin.

Kaydı QuickTime (Mac: ⌘⇧5) ya da telefon ekran kaydıyla alıp tek videoda birleştirin; Meta 1 GB'a kadar MP4/MOV
kabul eder. Videoda gerçek müşteri verisi görünmesin (deneme hesabı kullanın).

### Meta test hesabı isterse

Meta bazen inceleme için giriş bilgisi ister. O durumda `USERS_JSON`'a yalnızca Mesajlar'ı görebilen bir
`metareview` kullanıcısı ekleyip (MESAJ_USERNAMES'e dahil) şifresini inceleme formuna yazın; inceleme bitince
kullanıcıyı kaldırın.

### Sık ret sebepleri

- Ekran kaydında iznin gerçekten kullanıldığı görünmüyor (DM'nin gelişi + yanıt mutlaka görünmeli).
- Gizlilik politikası adresi açılmıyor ya da mesajlaşma verisinden bahsetmiyor (sayfa güncellendi).
- İşletme doğrulaması tamamlanmadan gönderilen başvuru bekletilir.
