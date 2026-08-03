import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import SmsPanel from "@/components/SmsPanel";

export const dynamic = "force-dynamic";

export default async function SmsPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/sms");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 260 }}>
          <h1>SMS Gönder</h1>
          <p className="subtitle">
            Müşterilere kargo, sipariş ve bilgilendirme mesajı gönderin.
            Gönderilen her mesaj geçmişe kaydedilir.
          </p>
        </div>
        <Link href="/panel" className="btn small secondary">
          ← Sipariş Paneli
        </Link>
      </div>

      <SmsPanel />

      <p style={{ color: "var(--muted)", fontSize: 13 }}>
        Kampanya, tanıtım ve kutlama mesajları <strong>ticari elektronik ileti</strong>{" "}
        sayılır; İYS (İleti Yönetim Sistemi) onayı olmayan numaralara
        gönderilmesi yasaktır. Kargo ve sipariş bilgilendirmeleri bu kapsamda
        değildir.
      </p>
    </main>
  );
}
