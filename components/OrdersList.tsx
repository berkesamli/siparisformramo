"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import {
  orderBalance,
  PAYMENT_LABELS,
  type SavedOrder,
  type OrderStatus,
} from "@/lib/orders";
import TahsilatModal, { type TahsilatBaglam } from "./TahsilatModal";

const STATUS_LABELS: Record<OrderStatus, string> = {
  olusturuldu: "Oluşturuldu",
  hazirlaniyor: "Hazırlanıyor",
  tamamlandi: "Tamamlandı",
};

const STATUS_COLORS: Record<OrderStatus, string> = {
  olusturuldu: "az",
  hazirlaniyor: "var",
  tamamlandi: "var",
};

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export default function OrdersList({ eldenSatis = false }: { eldenSatis?: boolean }) {
  const [filter, setFilter] = useState<{ range?: string; date?: string }>({
    range: "today",
  });
  const [orders, setOrders] = useState<SavedOrder[] | null>(null);
  const [error, setError] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>("all");

  const load = useCallback(async () => {
    setOrders(null);
    setError("");
    const qs = filter.date
      ? `date=${filter.date}`
      : `range=${filter.range || "today"}`;
    try {
      const res = await fetch(`/api/orders?${qs}`);
      const data = await res.json();
      if (data.ok) setOrders(data.orders);
      else setError(data.error || "Siparişler alınamadı.");
    } catch {
      setError("Sunucuya ulaşılamadı.");
    }
  }, [filter]);

  useEffect(() => {
    load();
  }, [load]);

  async function changeStatus(o: SavedOrder, status: OrderStatus) {
    // iyimser güncelleme
    setOrders((os) =>
      (os || []).map((x) => (x.orderId === o.orderId ? { ...x, status } : x))
    );
    const res = await fetch(
      `/api/orders/one?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ status }),
      }
    ).catch(() => null);
    if (!res || !res.ok) {
      setError("Durum güncellenemedi, sayfayı yenileyin.");
      load();
    }
  }

  // Ödeme girişi artık tahsilat kaydı üretir (tarih/yöntem/şube ile) —
  // pay-select'in yerini TahsilatModal aldı.
  const [tahsilatBaglam, setTahsilatBaglam] = useState<TahsilatBaglam | null>(null);

  const visible = (orders || []).filter(
    (o) => statusFilter === "all" || o.status === statusFilter
  );

  return (
    <div className="card">
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 16 }}>
        <button
          className={`btn small ${filter.range === "today" && !filter.date ? "" : "secondary"}`}
          onClick={() => setFilter({ range: "today" })}
        >
          Bugün
        </button>
        <button
          className={`btn small ${filter.range === "week" ? "" : "secondary"}`}
          onClick={() => setFilter({ range: "week" })}
        >
          Son 7 Gün
        </button>
        <input
          type="date"
          style={{ width: "auto" }}
          value={filter.date || ""}
          onChange={(e) =>
            e.target.value
              ? setFilter({ date: e.target.value })
              : setFilter({ range: "today" })
          }
        />
        <span style={{ flex: 1 }} />
        <select
          style={{ width: "auto" }}
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
        >
          <option value="all">Tüm Durumlar</option>
          <option value="olusturuldu">Oluşturuldu</option>
          <option value="hazirlaniyor">Hazırlanıyor</option>
          <option value="tamamlandi">Tamamlandı</option>
        </select>
        <button className="btn small secondary" onClick={load}>
          ↻ Yenile
        </button>
        {eldenSatis && (
          <button
            className="btn small"
            title="Ayaküstü perakende / teknik malzeme satışı — siparişsiz kasa girişi"
            onClick={() =>
              setTahsilatBaglam({ customerName: "PERAKENDE", serbest: true })
            }
          >
            💰 Elden Satış
          </button>
        )}
      </div>

      {error && <div className="notice err">{error}</div>}
      {!orders && !error && <p style={{ color: "var(--text-2)" }}>Yükleniyor…</p>}

      {orders && (
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Sipariş No</th>
                <th>Tarih</th>
                <th>Müşteri</th>
                <th>Çalışan</th>
                <th>Tutar</th>
                <th>Durum</th>
                <th>Ödeme</th>
                <th>İşlemler</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((o) => (
                <tr key={o.orderId}>
                  <td style={{ fontWeight: 600, whiteSpace: "nowrap" }}>{o.orderId}</td>
                  <td style={{ whiteSpace: "nowrap" }}>
                    {new Date(o.createdAt).toLocaleString("tr-TR", {
                      dateStyle: "short",
                      timeStyle: "short",
                      timeZone: "Europe/Istanbul",
                    })}
                  </td>
                  <td>{o.customer || "—"}</td>
                  <td>{o.employee}</td>
                  <td style={{ whiteSpace: "nowrap" }}>₺ {fmt(o.net)}</td>
                  <td>
                    <select
                      style={{ width: "auto", padding: "4px 8px", fontSize: 13 }}
                      value={o.status}
                      onChange={(e) => changeStatus(o, e.target.value as OrderStatus)}
                    >
                      {Object.entries(STATUS_LABELS).map(([k, v]) => (
                        <option key={k} value={k}>
                          {v}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td style={{ whiteSpace: "nowrap" }}>
                    <span className={`pay-select ${o.payment || "bekliyor"}`}
                      style={{ padding: "4px 8px", fontSize: 13, display: "inline-block" }}>
                      {PAYMENT_LABELS[o.payment || "bekliyor"]}
                    </span>{" "}
                    {orderBalance(o) > 0 && (
                      <button
                        className="btn small secondary"
                        title="Tahsilat gir"
                        onClick={() =>
                          setTahsilatBaglam({
                            customerId: o.customerId || undefined,
                            customerName: o.customer,
                            orderId: o.orderId,
                            orderDateKey: o.dateKey,
                            kalan: orderBalance(o),
                            branch: o.branch,
                          })
                        }
                      >
                        💰
                      </button>
                    )}
                    {o.payment === "kismi" && (
                      <div style={{ fontSize: 11, color: "var(--error)", marginTop: 2 }}>
                        Kalan ₺{fmt(orderBalance(o))}
                      </div>
                    )}
                  </td>
                  <td style={{ whiteSpace: "nowrap" }}>
                    <a
                      className="btn small secondary"
                      href={`/api/orders/pdf?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                      style={{ marginRight: 6 }}
                    >
                      ⬇ PDF
                    </a>
                    <Link
                      className="btn small secondary"
                      href={`/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                      style={{ marginRight: 6 }}
                    >
                      🖨️ Fiş
                    </Link>
                    <Link
                      className="btn small secondary"
                      href={`/panel/siparisler/duzenle?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                    >
                      ✏️ Düzenle
                    </Link>
                  </td>
                </tr>
              ))}
              {!visible.length && (
                <tr>
                  <td colSpan={7} style={{ color: "var(--muted)" }}>
                    Bu filtreye uyan sipariş yok.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {tahsilatBaglam && (
        <TahsilatModal
          baglam={tahsilatBaglam}
          onClose={() => setTahsilatBaglam(null)}
          onSaved={() => load()}
        />
      )}
    </div>
  );
}
