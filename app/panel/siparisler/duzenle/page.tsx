import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { getOrder } from "@/lib/orders";
import OrderForm, { type InitialOrder } from "@/components/OrderForm";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function OrderEditPage({
  searchParams,
}: {
  searchParams: { d?: string; id?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/siparisler");
  if (user.role !== "staff") redirect("/portal");

  const dateKey = searchParams.d || "";
  const orderId = searchParams.id || "";
  const order = dateKey && orderId ? await getOrder(dateKey, orderId) : null;

  const geriLink = (
    <Link href="/panel/siparisler" className="btn secondary">
      <Icon name="chevron-left" size={16} /> Siparişler
    </Link>
  );

  if (!order) {
    return (
      <main className="container">
        <PageHeader icon="edit" title="Sipariş Düzenle" actions={geriLink} />
        <div className="notice err">Sipariş bulunamadı.</div>
      </main>
    );
  }

  const initial: InitialOrder = {
    dateKey: order.dateKey,
    orderId: order.orderId,
    customer: order.customer,
    note: order.note,
    rate: order.rate,
    euroRate: order.euroRate,
    discountPct: order.discountPct,
    vatApplied: order.vatApplied,
    rows: (order.rows as InitialOrder["rows"]) || undefined,
  };

  return (
    <main className="container">
      <PageHeader
        icon="edit"
        title={<>Sipariş Düzenle — {order.orderId}</>}
        subtitle={
          <>Müşteri: {order.customer || "—"} · Oluşturan: {order.employee}</>
        }
        actions={geriLink}
      />

      {!order.rows?.length && (
        <div className="notice info">
          Bu sipariş eski sürümde kaydedildiği için satır detayları forma
          otomatik gelemedi — satırları yeniden girip kaydedin; sipariş
          numarası aynı kalır.
        </div>
      )}

      <OrderForm employeeName={user.name} initialOrder={initial} />
    </main>
  );
}
