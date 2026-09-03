import Link from "next/link";
import { redirect } from "next/navigation";
import { getDealerSession } from "@/lib/auth";
import { dealerCanOrder, getDealerPricing } from "@/lib/dealers";
import { getUsdRate } from "@/lib/kur";
import BayiWizard from "@/components/BayiWizard";

export const dynamic = "force-dynamic";

export default async function CercevePage() {
  const s = await getDealerSession();
  if (!s) redirect("/giris?next=/panel/cerceve");
  const [pricing, kur] = await Promise.all([getDealerPricing(s.dealer.slug), getUsdRate()]);

  return (
    <main className="container" style={{ maxWidth: 1200 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ marginBottom: 4 }}>Online Çerçeve</h1>
        <span style={{ flex: 1 }} />
        <Link href="/panel/siparisler" className="btn small secondary">📋 Siparişler</Link>
        <Link href="/panel/ayarlar" className="btn small secondary">⚙️ Fiyatlar</Link>
      </div>
      <p className="subtitle">Ölçü, çerçeve, paspartu, cam ve baskı seçin — fiyat sizin ayarlarınızla anında hesaplanır.</p>
      <BayiWizard
        dealer={{ name: s.dealer.name, phone: s.dealer.phone, website: s.dealer.website }}
        pricing={pricing}
        autoRate={kur?.usd ?? null}
        canOrder={dealerCanOrder(s.dealer)}
      />
    </main>
  );
}
