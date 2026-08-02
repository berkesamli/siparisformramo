import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import CustomerAccount from "@/components/CustomerAccount";

export const dynamic = "force-dynamic";

export default async function MusteriCariPage({
  searchParams,
}: {
  searchParams: { id?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/etiket");
  if (user.role !== "staff") redirect("/portal");

  const id = searchParams.id || "";
  if (!id) redirect("/etiket");

  return (
    <main className="container" style={{ maxWidth: 1100 }}>
      <CustomerAccount id={id} />
    </main>
  );
}
