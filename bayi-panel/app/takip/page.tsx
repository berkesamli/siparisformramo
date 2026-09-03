// Müşteri sipariş takip sayfası — giriş gerektirmez, imzalı link (k) ile açılır.
import { getDealer } from "@/lib/dealers";
import { getOrder, verifyTrackingToken } from "@/lib/orders";
import OrderReceipt from "@/components/OrderReceipt";

export const dynamic = "force-dynamic";

const STEPS = ["Beklemede", "Hazırlanıyor", "Hazır", "Teslim Edildi"] as const;

export default async function TakipPage({
  searchParams,
}: {
  searchParams: { b?: string; d?: string; id?: string; k?: string };
}) {
  const { b = "", d = "", id = "", k = "" } = searchParams;
  const valid = b && d && id && verifyTrackingToken(b, d, id, k);
  const dealer = valid ? await getDealer(b) : null;
  const order = dealer ? await getOrder(b, d, id) : null;

  if (!dealer || !order) {
    return (
      <main className="container" style={{ maxWidth: 560 }}>
        <div className="card" style={{ textAlign: "center", marginTop: 40 }}>
          <div style={{ fontSize: 40 }}>🔍</div>
          <h2>Sipariş bulunamadı</h2>
          <p style={{ color: "var(--muted)" }}>Takip linki hatalı ya da sipariş kaldırılmış olabilir.</p>
        </div>
      </main>
    );
  }

  const idx = STEPS.indexOf(order.status as (typeof STEPS)[number]);
  const pdfHref = `/api/siparisler/pdf?b=${b}&d=${d}&id=${encodeURIComponent(id)}&k=${k}`;

  return (
    <main className="container" style={{ maxWidth: 820 }}>
      <div className="card" style={{ marginTop: 20 }}>
        <h2 style={{ marginTop: 0 }}>Sipariş Durumu</h2>
        {order.status === "İptal" ? (
          <div className="notice err">Bu sipariş iptal edilmiştir.</div>
        ) : (
          <div className="rw-steps" style={{ marginBottom: 0 }}>
            {STEPS.map((s, i) => (
              <div key={s} className={`rw-step ${i === idx ? "active" : ""} ${i < idx ? "done" : ""}`}>
                <span className="rw-step-no">{i < idx ? "✓" : i + 1}</span>
                <span className="rw-step-label">{s}</span>
              </div>
            ))}
          </div>
        )}
        <p style={{ fontSize: 13.5, color: "var(--text-2)", marginBottom: 0 }}>
          {order.deliveryDate && order.status !== "Teslim Edildi" && <>Tahmini teslim: <strong>{order.deliveryDate}</strong> · </>}
          Sorularınız için: <strong>{dealer.name}</strong> — {dealer.phone}
        </p>
        <div style={{ marginTop: 12 }}>
          <a href={pdfHref} className="btn small secondary">⬇ Sipariş PDF</a>
        </div>
      </div>
      <OrderReceipt
        order={order}
        dealer={{ name: dealer.name, phone: dealer.phone, website: dealer.website, address: dealer.address }}
        customerView
      />
      <p style={{ textAlign: "center", fontSize: 12, color: "var(--muted)", margin: "16px 0 30px" }}>
        Bu sayfa Olga Çerçeve bayi sistemi tarafından oluşturulmuştur.
      </p>
    </main>
  );
}
