"use client";

// Müşteri cari kartı: sipariş geçmişi, toplam ciro, tahsilat ve kalan bakiye.

import { useEffect, useState } from "react";
import Link from "next/link";
import { customerTitle, type Customer } from "@/lib/customers";
import type { CariEntry } from "@/app/api/musteriler/cari/route";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

interface Summary {
  orderCount: number;
  totalAmount: number;
  totalPaid: number;
  balance: number;
  lastOrderAt: string | null;
}

export default function CustomerAccount({ id }: { id: string }) {
  const [customer, setCustomer] = useState<Customer | null>(null);
  const [entries, setEntries] = useState<CariEntry[]>([]);
  const [summary, setSummary] = useState<Summary | null>(null);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");

  useEffect(() => {
    let alive = true;
    fetch(`/api/musteriler/cari?id=${encodeURIComponent(id)}`)
      .then(async (r) => {
        const d = await r.json();
        if (!alive) return;
        if (!r.ok) {
          setErr(d.error || "Kayıt getirilemedi");
        } else {
          setCustomer(d.customer);
          setEntries(d.entries || []);
          setSummary(d.summary);
        }
      })
      .catch(() => alive && setErr("Sunucuya ulaşılamadı"))
      .finally(() => alive && setLoading(false));
    return () => {
      alive = false;
    };
  }, [id]);

  if (loading) return <p style={{ color: "var(--muted)" }}>Yükleniyor...</p>;
  if (err) return <div className="notice err">{err}</div>;
  if (!customer) return null;

  return (
    <div>
      <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 8 }}>
        <h1 style={{ marginBottom: 0 }}>{customerTitle(customer)}</h1>
        <span className={`lbl-branch ${customer.branch}`}>
          {customer.branch === "istanbul" ? "İSTANBUL" : "ANKARA"}
        </span>
        <span style={{ flex: 1 }} />
        <Link href="/etiket" className="btn small secondary">
          ← Müşteriler
        </Link>
      </div>
      <p className="subtitle">
        {[
          customer.phone,
          customer.email,
          [customer.district, customer.city].filter(Boolean).join(" / "),
        ]
          .filter(Boolean)
          .join("  ·  ") || "İletişim bilgisi girilmemiş"}
      </p>

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
        <div className={`cari-card ${(summary?.balance ?? 0) > 0 ? "borc" : ""}`}>
          <span>Kalan Bakiye</span>
          <strong style={{ color: (summary?.balance ?? 0) > 0 ? "var(--error)" : "var(--success)" }}>
            ₺{fmt(summary?.balance ?? 0)}
          </strong>
        </div>
      </div>

      <h2>Sipariş Geçmişi</h2>
      {entries.length === 0 ? (
        <div className="card" style={{ textAlign: "center", color: "var(--muted)" }}>
          Bu müşteriye ait sipariş bulunamadı.
          <br />
          <span style={{ fontSize: 12.5 }}>
            Sipariş alırken müşteri adını listeden seçerseniz kayıtlar buraya düşer.
          </span>
        </div>
      ) : (
        <div className="card" style={{ padding: 0, overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Sipariş</th>
                <th>Tür</th>
                <th>Tarih</th>
                <th>Durum</th>
                <th style={{ textAlign: "right" }}>Tutar</th>
                <th style={{ textAlign: "right" }}>Tahsilat</th>
                <th style={{ textAlign: "right" }}>Bakiye</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {entries.map((e) => (
                <tr key={`${e.kind}-${e.orderId}`}>
                  <td style={{ fontWeight: 700, color: "var(--brand)" }}>{e.orderId}</td>
                  <td>
                    <span className={`badge ${e.kind === "toptan" ? "var" : "az"}`}>
                      {e.kind === "toptan" ? "Toptan" : "Perakende"}
                    </span>
                  </td>
                  <td>{new Date(e.createdAt).toLocaleDateString("tr-TR")}</td>
                  <td style={{ fontSize: 12.5 }}>{e.status}</td>
                  <td style={{ textAlign: "right" }}>₺{fmt(e.total)}</td>
                  <td style={{ textAlign: "right", color: "var(--success)" }}>₺{fmt(e.paid)}</td>
                  <td style={{ textAlign: "right", fontWeight: 700, color: e.balance > 0 ? "var(--error)" : "var(--success)" }}>
                    ₺{fmt(e.balance)}
                  </td>
                  <td>
                    <Link
                      className="btn small secondary"
                      href={
                        e.kind === "toptan"
                          ? `/panel/siparisler/detay?d=${e.dateKey}&id=${encodeURIComponent(e.orderId)}`
                          : `/panel/perakende/siparisler/detay?d=${e.dateKey}&id=${encodeURIComponent(e.orderId)}`
                      }
                    >
                      Fiş
                    </Link>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
