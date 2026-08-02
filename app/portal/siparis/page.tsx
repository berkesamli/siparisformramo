import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import DealerRequestForm from "@/components/DealerRequestForm";

export const dynamic = "force-dynamic";

export default async function BayiSiparisPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/portal/siparis");

  return (
    <main className="container" style={{ maxWidth: 1000 }}>
      <h1>Sipariş Talebi</h1>
      <p className="subtitle">
        Almak istediğiniz ürünleri bildirin — ekibimiz fiyat teyidiyle size dönüş yapar.
      </p>
      <DealerRequestForm userName={user.name} />
    </main>
  );
}
