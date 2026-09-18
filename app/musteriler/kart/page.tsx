import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { finansAktif } from "@/data/users";
import CustomerAccount from "@/components/CustomerAccount";

export const dynamic = "force-dynamic";

// Müşteri kartı: iletişim, hızlı işlemler, Mikro bakiyesi ve sipariş geçmişi.
export default async function MusteriKartPage({
  searchParams,
}: {
  searchParams: { id?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/musteriler");
  if (user.role !== "staff") redirect("/portal");

  const id = searchParams.id || "";
  if (!id) redirect("/musteriler");

  return (
    <main className="container" style={{ maxWidth: 1100 }}>
      {/* Tahsilatlar bu sistemde işlenmediği sürece bakiye yalnızca Mikro'dan
          gösterilir; iç tahsilat/bakiye kutuları finans modülüyle (FINANS_AKTIF=1) açılır. */}
      <CustomerAccount id={id} finans={finansAktif()} />
    </main>
  );
}
