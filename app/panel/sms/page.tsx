import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import SmsPanel from "@/components/SmsPanel";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function SmsPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/sms");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <PageHeader
        title="SMS Gönder"
        subtitle="Müşterilere kargo, sipariş ve bilgilendirme mesajı gönderin. Gönderilen her mesaj geçmişe kaydedilir."
        icon="message"
        actions={
          <Link href="/panel" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Sipariş Paneli
          </Link>
        }
      />

      <SmsPanel />

      <p className="muted" style={{ fontSize: 13, marginTop: 16 }}>
        Kampanya, tanıtım ve kutlama mesajları <strong>ticari elektronik ileti</strong>{" "}
        sayılır; İYS (İleti Yönetim Sistemi) onayı olmayan numaralara
        gönderilmesi yasaktır. Kargo ve sipariş bilgilendirmeleri bu kapsamda
        değildir.
      </p>
    </main>
  );
}
