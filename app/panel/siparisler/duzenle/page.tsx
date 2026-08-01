import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import { getOrder } from "@/lib/orders";
import OrderForm, { type InitialOrder } from "@/components/OrderForm";

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

  if (!order) {
    return (
      <main className="container">
        <div className="notice err">Sipariş bulunamadı.</div>
        <Link href="/panel/siparisler" className="btn small secondary">
          ← Siparişler
        </Link>
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
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ flex: 1 }}>
          <h1>Sipariş Düzenle — {order.orderId}</h1>
          <p className="subtitle">
            Müşteri: {order.customer || "—"} · Oluşturan: {order.employee}
          </p>
        </div>
        <Link href="/panel/siparisler" className="btn small secondary">
          ← Siparişler
        </Link>
      </div>

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
