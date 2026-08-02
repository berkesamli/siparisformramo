import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import OrderForm from "@/components/OrderForm";
import AiChat from "@/components/AiChat";

export default async function PanelPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ marginBottom: 4 }}>Sipariş Paneli</h1>
        <span style={{ flex: 1 }} />
        <Link href="/panel/siparisler" className="btn small secondary">
          📋 Siparişler
        </Link>
      </div>
      <p className="subtitle">
        Hoş geldin {user.name} — siparişler e-posta ve WhatsApp ile iletilir.
      </p>
      <OrderForm employeeName={user.name} />
      <AiChat />
    </main>
  );
}
