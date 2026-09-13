import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import PriceListBrowser from "@/components/PriceListBrowser";
import PrintButton from "@/components/PrintButton";
import PageHeader from "@/components/PageHeader";

export default async function PriceListPage({
  searchParams,
}: {
  searchParams?: { q?: string; liste?: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/portal/fiyat-listesi");

  return (
    <main className="container">
      <PageHeader
        title="Toptan Fiyat Listesi"
        subtitle="Çerçeve profilleri ve teknik malzemeler — tüm fiyatlar KDV hariçtir. Arama kutusunun altındaki kutulardan liste değiştirin."
        icon="tag"
        actions={<PrintButton />}
      />

      <PriceListBrowser
        initialQuery={searchParams?.q}
        initialList={searchParams?.liste === "teknik" ? "teknik" : undefined}
      />

      <p className="muted" style={{ fontSize: 13 }}>
        Fiyatlar güncellenebilir; en güncel fiyat için 0850 305 75 45 numaralı
        sipariş hattımızla iletişime geçiniz. Satış yapılırken fiyatlara KDV
        ilave edilir.
      </p>
    </main>
  );
}
