"use client";

// Müşteri cari kartı: sipariş geçmişi, tahsilat hareketleri, açılış bakiyesi,
// toplam ciro ve kalan bakiye. "Tahsilat Ekle" ile buradan ödeme girilir.

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import Icon from "@/components/shell/Icon";
import PageHeader from "@/components/PageHeader";
import { customerTitle, type Customer } from "@/lib/customers";
import type { CariEntry } from "@/app/api/musteriler/cari/route";
import { TAHSILAT_YONTEM_LABELS, type Tahsilat } from "@/lib/tahsilat";
import TahsilatModal from "./TahsilatModal";
import MikroCariKutusu from "./MikroCariKutusu";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

interface Summary {
  orderCount: number;
  totalAmount: number;
  totalPaid: number;
  openingBalance: number;
  openingAsOf: string | null;
  balance: number;
  lastOrderAt: string | null;
}

export default function CustomerAccount({ id }: { id: string }) {
  const [customer, setCustomer] = useState<Customer | null>(null);
  const [entries, setEntries] = useState<CariEntry[]>([]);
  const [movements, setMovements] = useState<Tahsilat[]>([]);
  const [summary, setSummary] = useState<Summary | null>(null);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [modalOpen, setModalOpen] = useState(false);

  const load = useCallback(() => {
    fetch(`/api/musteriler/cari?id=${encodeURIComponent(id)}`)
      .then(async (r) => {
        const d = await r.json();
        if (!r.ok) {
          setErr(d.error || "Kayıt getirilemedi");
        } else {
          setCustomer(d.customer);
          setEntries(d.entries || []);
          setMovements(d.movements || []);
          setSummary(d.summary);
          setErr("");
        }
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [id]);

  useEffect(() => {
    load();
  }, [load]);

  if (loading) return <div className="empty">Yükleniyor...</div>;
  if (err) return <div className="notice err">{err}</div>;
  if (!customer) return null;

  const bakiye = summary?.balance ?? 0;

  return (
    <div>
      <PageHeader
        icon="user"
        kicker="Cari Kart"
        title={
          <>
            {customerTitle(customer)}{" "}
            <span
              className={`lbl-branch ${customer.branch}`}
              style={{ verticalAlign: "middle", marginLeft: 6 }}
            >
              {customer.branch === "istanbul" ? "İSTANBUL" : "ANKARA"}
            </span>
          </>
        }
        subtitle={
          [
            customer.phone,
            customer.email,
            [customer.district, customer.city].filter(Boolean).join(" / "),
          ]
            .filter(Boolean)
            .join("  ·  ") || "İletişim bilgisi girilmemiş"
        }
        actions={
          <>
            <button className="btn small" onClick={() => setModalOpen(true)}>
              <Icon name="wallet" size={15} /> Tahsilat Ekle
            </button>
            <Link href="/etiket" className="btn small secondary">
              <Icon name="chevron-left" size={15} /> Müşteriler
            </Link>
          </>
        }
      />

      {/* Cari özet kutuları */}
      <div className="cari-cards">
        <div className="cari-card">
          <span>Sipariş Sayısı</span>
          <strong>{summary?.orderCount ?? 0}</strong>
        </div>
        <div className="cari-card">
          <span>Toplam Ciro</span>
          <strong>₺{fmt(summary?.totalAmount ?? 0)}</strong>
        </div>
        <div className="cari-card">
          <span>Tahsil Edilen</span>
          <strong style={{ color: "var(--success)" }}>₺{fmt(summary?.totalPaid ?? 0)}</strong>
        </div>
        <div className={`cari-card ${bakiye > 0 ? "borc" : ""}`}>
          <span>Kalan Bakiye</span>
          <strong style={{ color: bakiye > 0 ? "var(--error)" : "var(--success)" }}>
            ₺{fmt(bakiye)}
          </strong>
        </div>
      </div>

      {summary && summary.openingBalance !== 0 && (
        <p className="notice info">
          Devir (açılış) bakiyesi: <strong>₺{fmt(summary.openingBalance)}</strong>
          {summary.openingAsOf && <> — {summary.openingAsOf} tarihi itibarıyla, Excel&apos;den aktarıldı.</>}
        </p>
      )}

      {/* ---- Mikro (resmi) cari: eşleştirme + canlı bakiye ---- */}
      <MikroCariKutusu customerId={customer.id} />

      {/* ---- Tahsilat hareketleri ---- */}
      <div className="card" style={{ marginTop: 18 }}>
        <div className="card-head">
          <span className="card-head-icon"><Icon name="wallet" size={16} /></span>
          <div>
            <h2>Tahsilat Hareketleri</h2>
            <span className="card-head-sub">{movements.length} hareket</span>
          </div>
        </div>
        {movements.length === 0 ? (
          <div className="empty" style={{ padding: "18px 12px" }}>
            Henüz tahsilat kaydı yok. &quot;Tahsilat Ekle&quot; ile ödeme girebilirsiniz.
          </div>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Tarih</th>
                  <th>Yöntem</th>
                  <th>Şube</th>
                  <th className="num">Tutar</th>
                  <th>Sipariş</th>
                  <th>Kaydeden</th>
                  <th>Not</th>
                </tr>
              </thead>
              <tbody>
                {movements.map((t) => (
                  <tr key={t.id}>
                    <td style={{ whiteSpace: "nowrap" }}>{t.dateKey.split("-").reverse().join(".")}</td>
                    <td style={{ whiteSpace: "nowrap" }}>{TAHSILAT_YONTEM_LABELS[t.method] || t.method}</td>
                    <td style={{ fontSize: 12.5 }}>
                      {t.branch === "istanbul" ? "İST" : "ANK"}
                    </td>
                    <td className="num" style={{ color: "var(--success)", fontWeight: 600, whiteSpace: "nowrap" }}>
                      {t.currency === "TL" ? "₺" : t.currency === "USD" ? "$" : "€"}
                      {fmt(t.amount)}
                    </td>
                    <td style={{ fontSize: 12.5, whiteSpace: "nowrap" }}>{t.orderId || "—"}</td>
                    <td style={{ fontSize: 12.5 }}>{t.tahsilEden || t.createdBy}</td>
                    <td className="muted" style={{ fontSize: 12.5 }}>{t.note || ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* ---- Sipariş geçmişi ---- */}
      <div className="card">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="list" size={16} /></span>
          <div>
            <h2>Sipariş Geçmişi</h2>
            <span className="card-head-sub">{entries.length} sipariş</span>
          </div>
        </div>
        {entries.length === 0 ? (
          <div className="empty" style={{ padding: "18px 12px" }}>
            Bu müşteriye ait sipariş bulunamadı.
            <br />
            <span style={{ fontSize: 12.5 }}>
              Sipariş alırken müşteri adını listeden seçerseniz kayıtlar buraya düşer.
            </span>
          </div>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Sipariş</th>
                  <th>Tür</th>
                  <th>Tarih</th>
                  <th>Durum</th>
                  <th className="num">Tutar</th>
                  <th className="num">Tahsilat</th>
                  <th className="num">Bakiye</th>
                  <th className="no-print"></th>
                </tr>
              </thead>
              <tbody>
                {entries.map((e) => (
                  <tr key={`${e.kind}-${e.orderId}`}>
                    <td style={{ fontWeight: 700, color: "var(--brand)", whiteSpace: "nowrap" }}>{e.orderId}</td>
                    <td>
                      <span className={`badge ${e.kind === "toptan" ? "var" : "az"}`}>
                        {e.kind === "toptan" ? "Toptan" : "Perakende"}
                      </span>
                    </td>
                    <td style={{ whiteSpace: "nowrap" }}>{new Date(e.createdAt).toLocaleDateString("tr-TR")}</td>
                    <td style={{ fontSize: 12.5 }}>{e.status}</td>
                    <td className="num" style={{ whiteSpace: "nowrap" }}>₺{fmt(e.total)}</td>
                    <td className="num" style={{ color: "var(--success)", whiteSpace: "nowrap" }}>₺{fmt(e.paid)}</td>
                    <td
                      className="num"
                      style={{ fontWeight: 700, whiteSpace: "nowrap", color: e.balance > 0 ? "var(--error)" : "var(--success)" }}
                    >
                      ₺{fmt(e.balance)}
                    </td>
                    <td className="no-print">
                      <Link
                        className="btn small secondary"
                        href={
                          e.kind === "toptan"
                            ? `/panel/siparisler/detay?d=${e.dateKey}&id=${encodeURIComponent(e.orderId)}`
                            : `/panel/perakende/siparisler/detay?d=${e.dateKey}&id=${encodeURIComponent(e.orderId)}`
                        }
                      >
                        <Icon name="file-text" size={14} /> Fiş
                      </Link>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {modalOpen && (
        <TahsilatModal
          baglam={{
            customerId: customer.id,
            customerName: customerTitle(customer),
            kalan: summary?.balance,
            branch: customer.branch,
          }}
          onClose={() => setModalOpen(false)}
          onSaved={load}
        />
      )}
    </div>
  );
}
