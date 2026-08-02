import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import RequestsList from "@/components/RequestsList";

export const dynamic = "force-dynamic";

export default async function TaleplerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/talepler");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1000 }}>
      <h1>Bayi Talepleri</h1>
      <p className="subtitle">
        Müşteri portalından gelen sipariş talepleri — onaylayıp siparişe dönüştürün.
      </p>
      <RequestsList />
    </main>
  );
}
