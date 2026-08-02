"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import {
  orderBalance,
  PAYMENT_LABELS,
  type SavedOrder,
  type OrderStatus,
  type PaymentStatus,
} from "@/lib/orders";

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

export default function OrdersList() {
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

  async function changePayment(o: SavedOrder, payment: PaymentStatus) {
    let paidAmount: number | undefined;
    if (payment === "kismi") {
      const girilen = prompt(
        `Tahsil edilen tutar (toplam ₺${fmt(o.net)}):`,
        String(o.paidAmount || "")
      );
      if (girilen === null) return;
      paidAmount = Number(girilen.replace(",", ".")) || 0;
    }
    const optimistic: SavedOrder = {
      ...o,
      payment,
      paidAmount:
        payment === "odendi" ? o.net : payment === "bekliyor" ? 0 : paidAmount ?? 0,
    };
    setOrders((os) => (os || []).map((x) => (x.orderId === o.orderId ? optimistic : x)));

    const res = await fetch(
      `/api/orders/one?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(
          paidAmount !== undefined ? { paidAmount } : { payment }
        ),
      }
    ).catch(() => null);
    if (!res || !res.ok) {
      setError("Ödeme durumu güncellenemedi, sayfayı yenileyin.");
      load();
    }
  }

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
                  <td>
                    <select
                      className={`pay-select ${o.payment || "bekliyor"}`}
                      style={{ width: "auto", padding: "4px 8px", fontSize: 13 }}
                      value={o.payment || "bekliyor"}
                      onChange={(e) => changePayment(o, e.target.value as PaymentStatus)}
                    >
                      {Object.entries(PAYMENT_LABELS).map(([k, v]) => (
                        <option key={k} value={k}>
                          {v}
                        </option>
                      ))}
                    </select>
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
    </div>
  );
}
