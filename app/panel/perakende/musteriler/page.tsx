import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import RetailCustomerManager from "@/components/RetailCustomerManager";

export const dynamic = "force-dynamic";

export default async function PerakendeMusterilerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/perakende/musteriler");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container">
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ marginBottom: 4 }}>Perakende Müşterileri</h1>
        <span style={{ flex: 1 }} />
        <Link href="/panel/perakende" className="btn small secondary">
          🖼️ Yeni Sipariş
        </Link>
        <Link href="/panel/perakende/siparisler" className="btn small secondary">
          📋 Perakende Siparişler
        </Link>
      </div>
      <p className="subtitle">
        Mağaza müşterileri — etiket/toptan defterinden ayrıdır. Sipariş
        kaydedilince müşteri telefon numarasıyla buraya kendiliğinden işlenir.
      </p>
      <RetailCustomerManager />
    </main>
  );
}
