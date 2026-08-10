import { redirect } from "next/navigation";
import Link from "next/link";
import { getSessionUser } from "@/lib/auth";
import OrdersList from "@/components/OrdersList";

export const dynamic = "force-dynamic";

export default async function TamamlananSiparislerPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/siparisler/tamamlanan");
  if (user.role !== "staff") redirect("/portal");

  return (
    <main className="container" style={{ maxWidth: 1480 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ flex: 1 }}>
          <h1>Tamamlanan Siparişler</h1>
          <p className="subtitle">
            Durumu “Tamamlandı”, ödemesi alınmış ve kontrol edilmiş siparişler —
            aktif listeden çıkıp buraya düşer.
          </p>
        </div>
        <Link href="/panel/siparisler" className="btn small secondary">
          ← Aktif Siparişler
        </Link>
      </div>
      <OrdersList tamamlananlar />
    </main>
  );
}
