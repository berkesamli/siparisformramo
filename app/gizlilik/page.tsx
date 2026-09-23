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

      <h2 id="veri-silme" style={{ fontSize: 18, marginTop: 24 }}>Veri silme talebi</h2>
      <p>
        Instagram, WhatsApp ya da e-posta yoluyla bizimle yaptığınız yazışmaların ve size ait kayıtların
        silinmesini istiyorsanız <strong>olgacercevee@gmail.com</strong> adresine e-posta gönderin ya da
        <strong> 0850 305 75 45</strong> numarasını arayın. Talebinizde Instagram kullanıcı adınızı, telefon
        numaranızı ya da e-posta adresinizi belirtmeniz yeterlidir. Kayıtlar en geç 30 gün içinde silinir ve
        size bilgi verilir. Yasal saklama yükümlülüğü bulunan sipariş ve fatura kayıtları bu sürenin dışındadır.
      </p>

      <h2 style={{ fontSize: 18, marginTop: 24 }}>İletişim</h2>
      <p>
        Verilerinizle ilgili her türlü talep için: Olga Çerçeve · 0850 305 75 45 · olgacercevee@gmail.com
      </p>

      <hr style={{ margin: "32px 0", border: 0, borderTop: "1px solid var(--border, #ddd)" }} />

      <h2 id="english" style={{ fontSize: 18, marginTop: 24 }}>Privacy Policy (English summary)</h2>
      <p>
        Olga Çerçeve Sanayi ve Ticaret Ltd. Şti. (Ankara, Türkiye) operates this order-management system for its
        own picture-frame business. We process only the data needed to prepare and deliver orders and to answer
        customer messages: customer name, phone number, e-mail address, order details and the content of the
        conversations customers start with us.
      </p>
      <p>
        Messages sent to our Instagram professional account (@olga.cerceve), to our WhatsApp business number or
        to our e-mail addresses are received through the official Meta APIs and Gmail and shown in our internal
        inbox so that our staff can reply. Message content, sender name and attachments are stored only for the
        purpose of handling the conversation and the related order. Replies are written by our staff; an AI
        assistant may suggest a draft, but nothing is sent without a staff member&apos;s action. We do not sell
        or share this data with third parties; it is transmitted only to the messaging provider (Meta, Google, SMS
        provider) needed to deliver our reply.
      </p>

      <h2 id="data-deletion" style={{ fontSize: 18, marginTop: 24 }}>Data deletion instructions</h2>
      <p>
        To request deletion of your conversations and personal data, e-mail <strong>olgacercevee@gmail.com</strong>{" "}
        or call <strong>+90 850 305 75 45</strong> and state your Instagram username, phone number or e-mail
        address. We delete the records within 30 days and confirm by reply. Order and invoice records that we are
        legally required to keep are excluded.
      </p>
    </main>
  );
}
