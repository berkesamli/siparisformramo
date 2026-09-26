import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import OrderForm, { type InitialCustomer } from "@/components/OrderForm";
import { getCustomer, customerTitle } from "@/lib/customers";
import AiChat from "@/components/AiChat";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function PanelPage({
  searchParams,
}: {
  searchParams?: { musteri?: string; metin?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel");
  if (user.role !== "staff") redirect("/portal");

  // Müşteri kartındaki "Yeni Sipariş": /panel?musteri=C123 → müşteri seçili açılır
  let onMusteri: InitialCustomer | undefined;
  const musteriId = searchParams?.musteri || "";
  if (/^[A-Za-z0-9]{2,40}$/.test(musteriId)) {
    const c = await getCustomer(musteriId);
    if (c) onMusteri = { id: c.id, title: customerTitle(c), branch: c.branch, iskontoPct: c.iskontoPct };
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
      <OrderForm employeeName={user.name} initialCustomer={onMusteri} autoImport={searchParams?.metin === "1"} />
      <AiChat />
    </main>
  );
}
