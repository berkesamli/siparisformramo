"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import {
  orderBalance,
  siparisTamamlandi,
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


export default function OrdersList({
  eldenSatis = false,
  // Arşiv modu: yalnızca tamamlanmış siparişleri (durum + ödeme + kontrol)
  // listeler. Normal modda bu siparişler aktif listeden gizlenir.
  tamamlananlar = false,
}: {
  eldenSatis?: boolean;
  tamamlananlar?: boolean;
}) {
  const [filter, setFilter] = useState<{ range?: string; date?: string; q?: string }>({
    range: "today",
  });
  const [orders, setOrders] = useState<SavedOrder[] | null>(null);
  const [error, setError] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>("all");
  // Arama kutusu — yazılan metin, "Ara" ile filtreye taşınır (tüm geçmişte arar)
  const [aramaMetni, setAramaMetni] = useState("");
  // Mesai sonrası siparişler gözden kaçmasın: son 7 günün kontrol
  // edilmemiş sipariş sayısı, hangi filtre açık olursa olsun üstte görünür.
  const [kontrolsuzSayi, setKontrolsuzSayi] = useState<number | null>(null);
  const [sadeceKontrolsuz, setSadeceKontrolsuz] = useState(false);

  const refreshKontrolsuz = useCallback(async () => {
    try {
      const res = await fetch("/api/orders?range=week");
      const data = await res.json();
      if (data.ok) {
        setKontrolsuzSayi(
          (data.orders as SavedOrder[]).filter((o) => !o.kontrol).length
        );
      }
    } catch {
      /* bant gösterilmez, liste etkilenmez */
    }
  }, []);

  useEffect(() => {
    refreshKontrolsuz();
  }, [refreshKontrolsuz]);

  const load = useCallback(async () => {
    setOrders(null);
    setError("");
    const qs = filter.q
      ? `q=${encodeURIComponent(filter.q)}`
      : filter.date
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

  async function toggleKontrol(o: SavedOrder) {
    const yeni = !o.kontrol;
    // iyimser güncelleme — işaret anında görünsün
    setOrders((os) =>
      (os || []).map((x) =>
        x.orderId === o.orderId
          ? { ...x, kontrol: yeni ? { by: "…", at: new Date().toISOString() } : undefined }
          : x
      )
    );
    setKontrolsuzSayi((n) => (n === null ? n : Math.max(0, n + (yeni ? -1 : 1))));
    const res = await fetch(
      `/api/orders/one?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ kontrol: yeni }),
      }
    ).catch(() => null);
    if (!res || !res.ok) {
      setError("Kontrol işareti kaydedilemedi, sayfayı yenileyin.");
      load();
      refreshKontrolsuz();
    } else {
      // sunucunun yazdığı isim/zaman gelsin
      const d = await res.json().catch(() => null);
      if (d?.ok) {
        setOrders((os) =>
          (os || []).map((x) =>
            x.orderId === o.orderId
              ? { ...x, kontrol: d.kontrol || undefined }
              : x
          )
        );
      }
    }
  }

  // Ödeme girişi artık tahsilat kaydı üretir (tarih/yöntem/şube ile) —
  // pay-select'in yerini TahsilatModal aldı.
  const [tahsilatBaglam, setTahsilatBaglam] = useState<TahsilatBaglam | null>(null);

  const visible = (orders || []).filter(
    (o) =>
      // Tamamlananlar arşivde, diğerleri aktif listede
      siparisTamamlandi(o) === tamamlananlar &&
      (statusFilter === "all" || o.status === statusFilter) &&
      (!sadeceKontrolsuz || !o.kontrol)
  );
  // Aktif listede gizlenen tamamlanmış sipariş sayısı (arşive yönlendirme için)
  const arsivlenen = tamamlananlar
    ? 0
    : (orders || []).filter((o) => siparisTamamlandi(o)).length;

  return (
    <div className="card">
      {/* Mesai sonrası girilen siparişler ertesi sabah gözden kaçmasın:
          son 7 günün kontrol edilmemişleri her filtrede üstte uyarır. */}
      {(kontrolsuzSayi ?? 0) > 0 && !sadeceKontrolsuz && !tamamlananlar && (
        <div
          className="notice info"
          style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 14 }}
        >
          <span style={{ flex: 1 }}>
            🔔 Son 7 günde <b>kontrol edilmemiş {kontrolsuzSayi} sipariş</b> var
            — akşam 19:00&apos;dan sonra girilenler dahil.
          </span>
          <button
            className="btn small"
            onClick={() => {
              setFilter({ range: "week" });
              setStatusFilter("all");
              setSadeceKontrolsuz(true);
            }}
          >
            Göster
          </button>
        </div>
      )}
      {sadeceKontrolsuz && (
        <div
          className="notice info"
          style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 14 }}
        >
          <span style={{ flex: 1 }}>
            Yalnızca <b>kontrol edilmemiş</b> siparişler listeleniyor (son 7 gün).
            Her siparişi inceleyip ✔ ile işaretleyin.
          </span>
          <button
            className="btn small secondary"
            onClick={() => setSadeceKontrolsuz(false)}
          >
            Tümünü Göster
          </button>
        </div>
      )}
      {/* Arama — müşteri adı, sipariş no, çalışan veya not içinde, tüm geçmişte */}
      <form
        style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}
        onSubmit={(e) => {
          e.preventDefault();
          const q = aramaMetni.trim();
          setSadeceKontrolsuz(false);
          setStatusFilter("all");
          setFilter(q ? { q } : { range: "today" });
        }}
      >
        <input
          style={{ flex: 1, minWidth: 200 }}
          placeholder="Ara: müşteri adı / sipariş no / çalışan"
          value={aramaMetni}
          onChange={(e) => setAramaMetni(e.target.value)}
        />
        <button className="btn small" type="submit">
          🔍 Ara
        </button>
        {filter.q && (
          <button
            className="btn small secondary"
            type="button"
            onClick={() => {
              setAramaMetni("");
              setFilter({ range: "today" });
            }}
          >
            ✕ Aramayı Temizle
          </button>
        )}
      </form>
      {filter.q && (
        <div className="notice info" style={{ marginBottom: 12 }}>
          🔍 <b>&quot;{filter.q}&quot;</b> için tüm sipariş geçmişinde arama sonuçları
          {orders ? ` — ${visible.length} sipariş bulundu.` : "…"}
        </div>
      )}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 16 }}>
        <button
          className={`btn small ${filter.range === "today" && !filter.date && !filter.q ? "" : "secondary"}`}
          onClick={() => { setFilter({ range: "today" }); setSadeceKontrolsuz(false); setAramaMetni(""); }}
        >
          Bugün
        </button>
        <button
          className={`btn small ${filter.range === "yesterday" ? "" : "secondary"}`}
          onClick={() => { setFilter({ range: "yesterday" }); setSadeceKontrolsuz(false); setAramaMetni(""); }}
        >
          Dün
        </button>
        <button
          className={`btn small ${filter.range === "week" && !sadeceKontrolsuz ? "" : "secondary"}`}
          onClick={() => { setFilter({ range: "week" }); setSadeceKontrolsuz(false); setAramaMetni(""); }}
        >
          Son 7 Gün
        </button>
        <input
          type="date"
          style={{ width: "auto" }}
          value={filter.date || ""}
          onChange={(e) => {
            setSadeceKontrolsuz(false);
            setAramaMetni("");
            if (e.target.value) setFilter({ date: e.target.value });
            else setFilter({ range: "today" });
          }}
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
        {!tamamlananlar && (
          <Link
            className="btn small secondary"
            href="/panel/siparisler/tamamlanan"
            title="Durumu tamamlandı, ödemesi alınmış ve kontrol edilmiş siparişler"
          >
            ✅ Tamamlananlar{arsivlenen > 0 ? ` (${arsivlenen})` : ""}
          </Link>
        )}
        {tamamlananlar && (
          <Link className="btn small secondary" href="/panel/siparisler">
            ← Aktif Siparişler
          </Link>
        )}
        {eldenSatis && !tamamlananlar && (
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
        <div className="ord-table-wrap">
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
                <th>Kontrol</th>
                <th className="ord-actions">İşlemler</th>
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
                      className={`status-select ${o.status}`}
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
                    {o.kontrol ? (
                      <button
                        className="btn small secondary"
                        style={{ color: "#15803d", borderColor: "#bbe3c8" }}
                        title={`${o.kontrol.by} kontrol etti — ${new Date(o.kontrol.at).toLocaleString("tr-TR", { dateStyle: "short", timeStyle: "short", timeZone: "Europe/Istanbul" })}. Geri almak için tıklayın.`}
                        onClick={() => toggleKontrol(o)}
                      >
                        ✔ {o.kontrol.by.split(" ")[0]}
                      </button>
                    ) : (
                      <button
                        className="btn small secondary"
                        title="Siparişi inceledikten sonra işaretleyin"
                        onClick={() => toggleKontrol(o)}
                      >
                        Kontrol Et
                      </button>
                    )}
                  </td>
                  {/* İşlemler sütunu sağa sabitlenir: tablo kaydırılsa bile
                      PDF / Fiş / Düzenle her zaman ekranda kalır. */}
                  <td className="ord-actions">
                    <a
                      className="btn small secondary"
                      href={`/api/orders/pdf?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                      title="Sipariş fişini PDF olarak indir"
                    >
                      ⬇ PDF
                    </a>
                    <Link
                      className="btn small secondary"
                      href={`/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                      title="Fişi görüntüle / yazdır"
                    >
                      🖨️ Fiş
                    </Link>
                    <Link
                      className="btn small secondary"
                      href={`/panel/siparisler/duzenle?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                      title="Siparişi düzenle"
                    >
                      ✏️ Düzenle
                    </Link>
                  </td>
                </tr>
              ))}
              {!visible.length && (
                <tr>
                  <td colSpan={9} style={{ color: "var(--muted)" }}>
                    {sadeceKontrolsuz
                      ? "🎉 Son 7 günün tüm siparişleri kontrol edildi."
                      : tamamlananlar
                        ? "Bu aralıkta tamamlanmış sipariş yok. Bir siparişin buraya düşmesi için durumu “Tamamlandı”, ödemesi alınmış ve kontrol edilmiş olmalı."
                        : "Bu filtreye uyan sipariş yok."}
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
