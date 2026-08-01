import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import RetailOrdersList from "@/components/RetailOrdersList";

export default async function PerakendeSiparislerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/perakende/siparisler");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1100 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ marginBottom: 4 }}>Perakende Siparişler</h1>
        <span style={{ flex: 1 }} />
        <Link href="/panel/perakende" className="btn small">
          ➕ Yeni Perakende Sipariş
        </Link>
      </div>
      <p className="subtitle">
        Çerçeveletme siparişlerini görüntüleyin ve durumlarını güncelleyin.
      </p>
      <RetailOrdersList />
    </main>
  );
}
