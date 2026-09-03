import { redirect } from "next/navigation";
import Link from "next/link";
import { getDealerSession } from "@/lib/auth";
import { getOrder, trackingPath } from "@/lib/orders";
import PrintButton from "@/components/PrintButton";
import OrderReceipt from "@/components/OrderReceipt";

export const dynamic = "force-dynamic";

export default async function OrderDetailPage({ searchParams }: { searchParams: { d?: string; id?: string } }) {
  const s = await getDealerSession();
  if (!s) redirect("/giris?next=/panel/siparisler");
  const order = searchParams.d && searchParams.id ? await getOrder(s.dealer.slug, searchParams.d, searchParams.id) : null;

  if (!order) {
    return (
      <main className="container">
        <div className="notice err">Sipariş bulunamadı.</div>
        <Link href="/panel/siparisler" className="btn small secondary">← Siparişler</Link>
      </main>
    );
  }

  const pdfHref = `/api/siparisler/pdf?d=${order.dateKey}&id=${encodeURIComponent(order.orderId)}`;

  return (
    <main className="container" style={{ maxWidth: 820 }}>
      <div className="no-print" style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <Link href="/panel/siparisler" className="btn small secondary">← Siparişler</Link>
        <span style={{ flex: 1 }} />
        <a href={trackingPath(order)} target="_blank" rel="noreferrer" className="btn small secondary">🔗 Müşteri Takip Sayfası</a>
        <a href={pdfHref} className="btn small secondary">⬇ Üretim PDF</a>
        <PrintButton />
      </div>
      <OrderReceipt order={order} dealer={{ name: s.dealer.name, phone: s.dealer.phone, website: s.dealer.website, address: s.dealer.address }} />
    </main>
  );
}
