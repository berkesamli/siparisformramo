import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import RetailWizard from "@/components/RetailWizard";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export default async function PerakendePage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/perakende");
  if (user.role !== "staff") redirect("/portal");

  return (
    // Genişlik globaldeki .container kurallarından gelir — geniş ekranda
    // sihirbaz + önizleme daha ferah yerleşir
    <main className="container">
      <PageHeader
        icon="frame"
        title="Online Çerçeve — Perakende"
        subtitle="Ölçü, çerçeve, paspartu, cam ve baskı seçin — fiyat anında hesaplanır."
        actions={
          <>
            <Link href="/panel/perakende/musteriler" className="btn secondary">
              <Icon name="users" size={16} /> Müşteriler
            </Link>
            <Link href="/panel/perakende/siparisler" className="btn secondary">
              <Icon name="image" size={16} /> Perakende Siparişler
            </Link>
          </>
        }
      />
      <RetailWizard employeeName={user.name} />
    </main>
  );
}
