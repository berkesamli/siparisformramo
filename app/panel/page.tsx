import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { getOrder } from "@/lib/orders";
import OrderForm, { type KopyaOrder } from "@/components/OrderForm";
import AiChat from "@/components/AiChat";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function PanelPage({
  searchParams,
}: {
  searchParams?: { kopya?: string; d?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel");
  if (user.role !== "staff") redirect("/portal");

  // "Bu siparişi kopyala": /panel?kopya=OLG-2026-042&d=2026-09-10
  // Eski sipariş okunur, satırları yeni sipariş taslağı olarak forma verilir.
  let kopya: KopyaOrder | undefined;
  const kopyaNo = searchParams?.kopya || "";
  const kopyaGun = searchParams?.d || "";
  if (kopyaNo && /^\d{4}-\d{2}-\d{2}$/.test(kopyaGun)) {
    const eski = await getOrder(kopyaGun, kopyaNo);
    if (eski) {
      kopya = {
        kaynakNo: eski.orderId,
        customer: eski.customer,
        customerId: eski.customerId || undefined,
        branch: eski.branch,
        discountPct: eski.discountPct || undefined,
        vatApplied: eski.vatApplied,
        rows: Array.isArray(eski.rows) ? (eski.rows as KopyaOrder["rows"]) : undefined,
      };
    }
  }

  return (
    <main className="container">
      <PageHeader
        icon="plus"
        title="Sipariş Paneli"
        subtitle={
          <>Hoş geldin {user.name} — siparişler e-posta ve WhatsApp ile iletilir.</>
        }
        actions={
          <Link href="/panel/siparisler" className="btn secondary">
            <Icon name="list" size={16} /> Siparişler
          </Link>
        }
      />
      <OrderForm employeeName={user.name} kopyaOrder={kopya} />
      <AiChat />
    </main>
  );
}
