// Uygulama içi duyurular ("Duyurular") — yeni özellik / bilgilendirme kartları.
//
// YENİ DUYURU EKLEMEK İÇİN aşağıdaki DUYURULAR listesine yeni bir kayıt ekleyin:
//   - id: benzersiz ve KALICI olmalı (ör. "ozellik-adi-2026-09"). Kullanıcının
//     duyuruyu gördüğü / "daha sonra" dediği bilgisi tarayıcıda bu id ile (kullanıcı
//     adına göre) hatırlanır; id değişirse duyuru herkese yeniden gösterilir.
//   - baslangic / bitis: YYYY-MM-DD (İstanbul günü). bitis dahildir; verilmezse
//     duyuru süresiz kalır. Tarih aralığı dışındaki duyurular hiç görünmez.
//   - roller: kime gösterileceği ("staff" = çalışan paneli, "customer" = müşteri portalı).
//   - href / hrefLabel: isteğe bağlı birincil bağlantı ve düğme yazısı.
//   - ikon: components/shell/Icon.tsx içindeki IconName; verilmezse "sparkles".
// Aktif olan ilk (en yeni) görülmemiş duyuru girişte bir kez modal olarak açılır.
// "Daha sonra" denirse modal bir daha açılmaz; zildeki kırmızı nokta ise kullanıcı
// zili açana dek kalır. Tüm aktif duyurular üst çubuktaki bildirim zilinin
// "Duyurular" bölümünde listelenir (zil yalnızca çalışan panelinde vardır).

import type { IconName } from "@/components/shell/Icon";

export interface Duyuru {
  id: string;
  baslik: string;
  metin: string;
  href?: string;
  hrefLabel?: string;
  baslangic: string; // YYYY-MM-DD, İstanbul
  bitis?: string;    // YYYY-MM-DD, dahil
  roller: ("staff" | "customer")[];
  ikon?: IconName;
}

export const DUYURULAR: Duyuru[] = [
  {
    id: "satislarim-2026-09",
    baslik: "Yeni: Satışlarım ekranı",
    metin:
      "Artık kendi cironuzu takip edebilirsiniz. Bugünkü, son 7 günlük ve aylık satışlarınızı, 14 günlük grafiği ve kendi siparişlerinizi 'Satışlarım' ekranında görebilirsiniz. Sadece sizin adınıza girilen siparişler sayılır.",
    href: "/panel/satislarim",
    hrefLabel: "Satışlarımı Gör",
    baslangic: "2026-09-15",
    bitis: "2026-10-15",
    roller: ["staff"],
    ikon: "trending-up",
  },
];
