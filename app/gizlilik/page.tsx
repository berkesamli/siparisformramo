// Gizlilik politikası — Meta uygulamasının "Canlı" moda alınması için
// herkese açık bir gizlilik sayfası gerekir. Giriş gerektirmez.

export const metadata = { title: "Gizlilik Politikası · Olga Çerçeve" };

export default function GizlilikPage() {
  return (
    <main style={{ maxWidth: 720, margin: "0 auto", padding: "40px 20px", lineHeight: 1.6 }}>
      <h1 style={{ fontSize: 26, marginBottom: 6 }}>Gizlilik Politikası</h1>
      <p style={{ color: "var(--muted)", marginBottom: 24 }}>Olga Çerçeve Sanayi ve Ticaret Ltd. Şti. · Sipariş sistemi</p>

      <h2 style={{ fontSize: 18, marginTop: 24 }}>Hangi veriler işlenir?</h2>
      <p>
        Sipariş sistemi, yalnızca siparişlerin hazırlanması ve teslimi için gerekli bilgileri işler:
        müşteri adı, telefon numarası, sipariş içeriği ve tutarı. Bu bilgiler müşterinin sipariş
        verirken kendisinin ilettiği bilgilerdir.
      </p>

      <h2 style={{ fontSize: 18, marginTop: 24 }}>WhatsApp ve SMS bildirimleri</h2>
      <p>
        Müşterinin telefon numarasına, yalnızca kendi siparişine ilişkin bilgilendirme mesajı
        (sipariş fişi) gönderilir. Bu mesajlar WhatsApp Business Platform ya da SMS aracılığıyla
        iletilir. Reklam veya kampanya mesajı gönderilmez. Bildirim almak istemeyen müşteriler
        0850 305 75 45 numaralı hattımızı arayarak numaralarının listeden çıkarılmasını isteyebilir.
      </p>

      <h2 style={{ fontSize: 18, marginTop: 24 }}>Instagram, WhatsApp ve e-posta yazışmaları</h2>
      <p>
        Instagram hesabımıza (@olga.cerceve), WhatsApp hattımıza ya da e-posta adreslerimize
        gönderdiğiniz mesajlar, size yanıt verebilmemiz için sipariş sistemimizin gelen kutusuna
        aktarılır. Mesaj içeriği, gönderen adı ve varsa ekler yalnızca yazışmayı yürütmek ve
        siparişinizi hazırlamak amacıyla saklanır; mesajlar Instagram veya WhatsApp üzerinden
        Meta'nın resmi API'leri aracılığıyla alınır ve yanıtlanır. Yazışmalarınızın silinmesini
        istediğinizde aşağıdaki iletişim kanalından talep edebilirsiniz.
      </p>

      <h2 style={{ fontSize: 18, marginTop: 24 }}>Paylaşım ve saklama</h2>
      <p>
        Veriler üçüncü kişilerle paylaşılmaz, satılmaz. Mesaj iletimi için WhatsApp (Meta) ve SMS
        sağlayıcısına yalnızca alıcı numarası ve mesaj içeriği aktarılır. Sipariş kayıtları ticari
        defter yükümlülükleri süresince saklanır.
      </p>

      <h2 style={{ fontSize: 18, marginTop: 24 }}>İletişim</h2>
      <p>
        Verilerinizle ilgili her türlü talep için: Olga Çerçeve · 0850 305 75 45
      </p>
    </main>
  );
}
