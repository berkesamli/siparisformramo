import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import RetailWizard from "@/components/RetailWizard";

export default async function PerakendePage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/perakende");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ marginBottom: 4 }}>Online Çerçeve — Perakende</h1>
        <span style={{ flex: 1 }} />
        <Link href="/panel/perakende/siparisler" className="btn small secondary">
          📋 Perakende Siparişler
        </Link>
      </div>
      <p className="subtitle">
        Ölçü, çerçeve, paspartu, cam ve baskı seçin — fiyat anında hesaplanır.
      </p>
      <RetailWizard employeeName={user.name} />
    </main>
  );
}
