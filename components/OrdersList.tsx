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
import { sayi } from "@/lib/num";
import Icon from "@/components/shell/Icon";

const STATUS_LABELS: Record<OrderStatus, string> = {
  olusturuldu: "Oluşturuldu",
  hazirlaniyor: "Hazırlanıyor",
  tamamlandi: "Tamamlandı",
  iptal: "İptal",
};

// Ödeme rozeti — .badge renk sınıfı (bekliyor kırmızı, kısmi sarı, ödendi yeşil)
const PAY_BADGE: Record<string, string> = {
  bekliyor: "err",
  kismi: "warn",
  odendi: "ok",
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
  // Patrona WhatsApp ile fiş gönderme düğmeleri — yalnızca sahipler (sayfa geçirir)
  patronGonderim = false,
}: {
  eldenSatis?: boolean;
  tamamlananlar?: boolean;
  patronGonderim?: boolean;
}) {
  const [filter, setFilter] = useState<{ range?: string; date?: string; q?: string }>({
    range: "today",
  });
  const [orders, setOrders] = useState<SavedOrder[] | null>(null);
  const [error, setError] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>("all");
  // Çalışan filtresi — yüklü listedeki çalışanlardan seçilir ("Murat'ın bugün aldıkları")
  const [employeeFilter, setEmployeeFilter] = useState<string>("all");
  // Patrona WhatsApp ile fiş gönderimi (tek sipariş ya da listedekilerin tümü)
  const [waGonderiyor, setWaGonderiyor] = useState(false);
  const [waSonuc, setWaSonuc] = useState<
    { giden: number; toplam: number; sonuclar: { id: string; ok: boolean; yontem?: string; hata?: string }[] } | null
  >(null);
  const [waHata, setWaHata] = useState("");
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
        // İptal edilen sipariş kontrol beklemez
        setKontrolsuzSayi(
          (data.orders as SavedOrder[]).filter(
            (o) => !o.kontrol && o.status !== "iptal"
          ).length
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
    // İptal geri alınabilir ama ciddi bir işlem — önce onay iste
    if (status === "iptal") {
      const onay = confirm(
        `${o.orderId} — ${o.customer || "müşteri"} siparişi iptal edilsin mi?\n\n` +
          "İptal edilen sipariş silinmez ama ciroya, raporlara ve müşteri " +
          "bakiyesine dahil edilmez. Gerekirse durumu tekrar değiştirilebilir."
      );
      if (!onay) return;
    }
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

  // İade: müşteriye para iadesi — orijinal siparişe bağlı NEGATİF tahsilat
  // kaydı düşülür, siparişin ödenen tutarı azalır (kasa/cari doğru kalır).
  async function iadeYap(o: SavedOrder) {
    const odenen = o.payment === "odendi" ? o.net : Number(o.paidAmount) || 0;
    if (!(odenen > 0)) return;
    const giris = prompt(
      `${o.orderId} — ${o.customer || "müşteri"}\nÖdenen: ₺${fmt(odenen)}\n\nİade tutarı (₺):`,
      String(odenen)
    );
    if (giris === null) return;
    const tutar = sayi(giris);
    if (!(tutar > 0) || tutar > odenen + 0.01) {
      alert("Geçersiz tutar — iade, ödenenden fazla olamaz.");
      return;
    }
    const onay = confirm(
      `₺${fmt(tutar)} iade kaydedilsin mi?\n\nKasadan düşülür ve ${o.orderId} ` +
        "siparişine negatif tahsilat olarak işlenir. Bu işlem geri alınamaz."
    );
    if (!onay) return;
    const res = await fetch(
      `/api/orders/iade?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ amount: tutar }),
      }
    ).catch(() => null);
    const dta = res ? await res.json().catch(() => null) : null;
    if (!res || !res.ok || !dta?.ok) {
      setError(dta?.error || "İade kaydedilemedi, sayfayı yenileyin.");
    }
    load();
  }

  const visible = (orders || []).filter(
    (o) =>
      // Tamamlananlar arşivde, diğerleri aktif listede
      siparisTamamlandi(o) === tamamlananlar &&
      (statusFilter === "all" || o.status === statusFilter) &&
      (employeeFilter === "all" || o.employee === employeeFilter) &&
      // Kontrol bekleyenler görünümünde iptaller listelenmez
      (!sadeceKontrolsuz || (!o.kontrol && o.status !== "iptal"))
  );
  // Aktif listede gizlenen tamamlanmış sipariş sayısı (arşive yönlendirme için)
  const arsivlenen = tamamlananlar
    ? 0
    : (orders || []).filter((o) => siparisTamamlandi(o)).length;
  // Çalışan seçenekleri: yüklü listedeki (filtre öncesi) çalışanlar
  const calisanlar = Array.from(new Set((orders || []).map((o) => o.employee).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "tr")
  );
  const WA_EN_FAZLA = 30;

  // Fiş(ler)i patrona WhatsApp ile gönder — sipariş kaydında olduğu gibi PDF dosyası gider
  async function patronaGonder(list: SavedOrder[]) {
    if (!list.length || waGonderiyor) return;
    const ozet =
      list.length === 1
        ? `${list[0].orderId} (${list[0].customer || "—"}) fişi`
        : `Listedeki ${list.length} sipariş fişi`;
    if (!window.confirm(`${ozet} patrona WhatsApp ile PDF olarak gönderilecek. Devam edilsin mi?`)) return;
    setWaGonderiyor(true);
    setWaSonuc(null);
    setWaHata("");
    try {
      const res = await fetch("/api/whatsapp/patron", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ orders: list.map((o) => ({ d: o.dateKey, id: o.orderId })) }),
      });
      const d = await res.json().catch(() => null);
      if (d && typeof d.toplam === "number") setWaSonuc(d);
      else setWaHata(d?.error || "Gönderim başarısız.");
    } catch {
      setWaHata("Sunucuya ulaşılamadı.");
    } finally {
      setWaGonderiyor(false);
    }
  }

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
          <Icon name="search" size={15} /> Ara
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
            <Icon name="x" size={14} /> Aramayı Temizle
          </button>
        )}
      </form>
      {filter.q && (
        <div className="notice info" style={{ marginBottom: 12 }}>
          🔍 <b>&quot;{filter.q}&quot;</b> için tüm sipariş geçmişinde arama sonuçları
          {orders ? ` — ${visible.length} sipariş bulundu.` : "…"}
        </div>
      )}
      <div className="row" style={{ marginBottom: 16 }}>
        {/* Tarih aralığı — segmentli kontrol (.seg); aktif seçenek .active */}
        <div className="seg" role="group" aria-label="Tarih aralığı">
          <button
            type="button"
            className={filter.range === "today" && !filter.date && !filter.q ? "active" : ""}
            onClick={() => { setFilter({ range: "today" }); setSadeceKontrolsuz(false); setAramaMetni(""); }}
          >
            Bugün
          </button>
          <button
            type="button"
            className={filter.range === "yesterday" ? "active" : ""}
            onClick={() => { setFilter({ range: "yesterday" }); setSadeceKontrolsuz(false); setAramaMetni(""); }}
          >
            Dün
          </button>
          <button
            type="button"
            className={filter.range === "week" && !sadeceKontrolsuz ? "active" : ""}
            onClick={() => { setFilter({ range: "week" }); setSadeceKontrolsuz(false); setAramaMetni(""); }}
          >
            Son 7 Gün
          </button>
          <button
            type="button"
            className={filter.range === "days15" ? "active" : ""}
            onClick={() => { setFilter({ range: "days15" }); setSadeceKontrolsuz(false); setAramaMetni(""); }}
          >
            Son 15 Gün
          </button>
        </div>
        <input
          type="date"
          style={{ width: "auto", maxWidth: "100%" }}
          value={filter.date || ""}
          onChange={(e) => {
            setSadeceKontrolsuz(false);
            setAramaMetni("");
            if (e.target.value) setFilter({ date: e.target.value });
            else setFilter({ range: "today" });
          }}
        />
        {/* Durum filtresi + eylemler tek grup olarak sağa yaslanır; dar
            ekranda hep birlikte alt satıra iner (tek başına kalan kontrol olmaz). */}
        <div className="row" style={{ marginLeft: "auto" }}>
          {calisanlar.length > 1 && (
            <select
              style={{ width: "auto", minWidth: 150, maxWidth: "100%" }}
              value={employeeFilter}
              onChange={(e) => setEmployeeFilter(e.target.value)}
              title="Çalışana göre süz"
            >
              <option value="all">Tüm Çalışanlar</option>
              {calisanlar.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          )}
          <select
            // minWidth: en uzun seçenek + base.css'in 34px ok boşluğu; ok yazının üstüne binmesin
            style={{ width: "auto", minWidth: 172, maxWidth: "100%" }}
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
          >
            <option value="all">Tüm Durumlar</option>
            <option value="olusturuldu">Oluşturuldu</option>
            <option value="hazirlaniyor">Hazırlanıyor</option>
            <option value="tamamlandi">Tamamlandı</option>
            <option value="iptal">İptal</option>
          </select>
          <button className="btn small secondary" type="button" onClick={load}>
            <Icon name="refresh" size={14} /> Yenile
          </button>
          {patronGonderim && (
          <button
            className="btn small secondary"
            type="button"
            disabled={waGonderiyor || visible.length === 0 || visible.length > WA_EN_FAZLA}
            title={
              visible.length > WA_EN_FAZLA
                ? `Tek seferde en fazla ${WA_EN_FAZLA} fiş — tarihi ya da çalışanı daraltın`
                : "Listedeki siparişlerin fişlerini patrona WhatsApp ile PDF olarak gönder"
            }
            onClick={() => patronaGonder(visible)}
          >
            <Icon name="message" size={14} /> {waGonderiyor ? "Gönderiliyor…" : `Patrona Gönder (${visible.length})`}
          </button>
          )}
          {!tamamlananlar && (
            <Link
              className="btn small secondary"
              href="/panel/siparisler/tamamlanan"
              title="Durumu tamamlandı, ödemesi alınmış ve kontrol edilmiş siparişler"
            >
              <Icon name="check-circle" size={14} /> Tamamlananlar{arsivlenen > 0 ? ` (${arsivlenen})` : ""}
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
              <Icon name="wallet" size={14} /> Elden Satış
            </button>
          )}
        </div>
      </div>

      {error && <div className="notice err">{error}</div>}
      {waHata && (
        <div className="notice err" style={{ display: "flex", gap: 12, alignItems: "center" }}>
          <span style={{ flex: 1 }}>WhatsApp: {waHata}</span>
          <button className="btn small secondary" type="button" onClick={() => setWaHata("")}>Kapat</button>
        </div>
      )}
      {waSonuc && (
        <div className={`notice ${waSonuc.giden === waSonuc.toplam ? "ok" : waSonuc.giden > 0 ? "warn" : "err"}`}>
          <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
            <span style={{ flex: 1 }}>
              <b>{waSonuc.giden}/{waSonuc.toplam}</b> fiş patrona WhatsApp ile gönderildi
              {waSonuc.sonuclar.some((r) => r.ok && r.yontem === "serbest") && " (serbest belge mesajı)"}.
            </span>
            <button className="btn small secondary" type="button" onClick={() => setWaSonuc(null)}>Kapat</button>
          </div>
          {waSonuc.sonuclar.filter((r) => !r.ok).map((r) => (
            <div key={r.id} style={{ marginTop: 4, fontSize: 13 }}>{r.id}: {r.hata || "gönderilemedi"}</div>
          ))}
        </div>
      )}
      {!orders && !error && <p className="text-2">Yükleniyor…</p>}

      {orders && (
        <>
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
                  <td
                    style={{
                      whiteSpace: "nowrap",
                      // İptalde tutar üstü çizili — ciroya girmediği belli olsun
                      textDecoration: o.status === "iptal" ? "line-through" : undefined,
                      color: o.status === "iptal" ? "var(--muted)" : undefined,
                    }}
                  >
                    ₺ {fmt(o.net)}
                  </td>
                  <td>
                    <select
                      className={`status-select ${o.status}`}
                      style={{ width: "auto", maxWidth: "100%", padding: "5px 30px 5px 10px", fontSize: 13 }}
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
                    <span className={`badge ${PAY_BADGE[o.payment || "bekliyor"] || ""} pay-select ${o.payment || "bekliyor"}`}>
                      {PAYMENT_LABELS[o.payment || "bekliyor"]}
                    </span>{" "}
                    {orderBalance(o) > 0 && (
                      <button
                        className="btn small secondary icon"
                        title="Tahsilat gir"
                        aria-label="Tahsilat gir"
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
                        <Icon name="wallet" size={15} />
                      </button>
                    )}
                    {(o.payment === "odendi" || (Number(o.paidAmount) || 0) > 0) && (
                      <button
                        className="btn small secondary icon"
                        title="İade — müşteriye para iadesi (siparişe bağlı negatif tahsilat kaydı düşülür)"
                        aria-label="İade"
                        onClick={() => iadeYap(o)}
                      >
                        ↩
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
                        style={{ color: "var(--success)", borderColor: "var(--success-line)", background: "var(--success-soft)" }}
                        title={`${o.kontrol.by} kontrol etti — ${new Date(o.kontrol.at).toLocaleString("tr-TR", { dateStyle: "short", timeStyle: "short", timeZone: "Europe/Istanbul" })}. Geri almak için tıklayın.`}
                        onClick={() => toggleKontrol(o)}
                      >
                        <Icon name="check-circle" size={14} /> {o.kontrol.by.split(" ")[0]}
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
                      <Icon name="download" size={14} /> PDF
                    </a>
                    {patronGonderim && (
                      <button
                        type="button"
                        className="btn small secondary icon"
                        title="Fişi patrona WhatsApp ile PDF olarak gönder"
                        aria-label="Fişi patrona WhatsApp ile gönder"
                        disabled={waGonderiyor}
                        onClick={() => patronaGonder([o])}
                      >
                        <Icon name="message" size={15} />
                      </button>
                    )}
                    <Link
                      className="btn small secondary"
                      href={`/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                      title="Fişi görüntüle / yazdır"
                    >
                      <Icon name="printer" size={14} /> Fiş
                    </Link>
                    <Link
                      className="btn small secondary"
                      href={`/panel/siparisler/duzenle?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`}
                      title="Siparişi düzenle"
                    >
                      <Icon name="edit" size={14} /> Düzenle
                    </Link>
                    <Link
                      className="btn small secondary"
                      href={`/panel?kopya=${encodeURIComponent(o.orderId)}&d=${o.dateKey}`}
                      title="Aynı satırlarla yeni sipariş aç — fiyatlar bugünün katalog fiyatı ve kurundan hesaplanır"
                    >
                      <Icon name="copy" size={14} /> Kopyala
                    </Link>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {/* Boş durum tablonun DIŞINDA: 9 sütunlu tablo telefonda yatay kayar,
            mesaj kart genişliğine göre ortalanır ve her ekranda okunur. */}
        {!visible.length && (
          <div className="empty">
            {sadeceKontrolsuz
              ? "🎉 Son 7 günün tüm siparişleri kontrol edildi."
              : tamamlananlar
                ? "Bu aralıkta tamamlanmış sipariş yok. Bir siparişin buraya düşmesi için durumu “Tamamlandı” olmalı ve kontrol edilmiş olmalı."
                : "Bu filtreye uyan sipariş yok."}
          </div>
        )}
        </>
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
