"use client";

// Müşteri kartı: iletişim çipleri (ara / WhatsApp / e-posta), hızlı işlemler
// (yeni sipariş, kargo etiketi, düzenle), Mikro'daki resmi bakiye ve sipariş
// geçmişi. finans=true (FINANS_AKTIF=1) olduğunda sistemin kendi tahsilat
// hareketleri, açılış bakiyesi ve kalan bakiye kutuları da görünür; kapalıyken
// tahsilatlar burada işlenmediği için bakiye yalnızca Mikro'dan okunur.

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import Icon from "@/components/shell/Icon";
import { customerTitle, musteriBolgesi, bolgeler as varsayilanBolgeler, type Bolge, type BolgeInfo, type Customer } from "@/lib/customers";
import type { CariEntry } from "@/app/api/musteriler/cari/route";
import { TAHSILAT_YONTEM_LABELS, type Tahsilat } from "@/lib/tahsilat";
import TahsilatModal from "./TahsilatModal";
import MikroCariKutusu from "./MikroCariKutusu";
import CustomerForm from "./CustomerForm";
import { initials } from "@/components/shell/nav-config";

const fmt = (n: number) =>
  (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const tarih = (iso: string) => new Date(iso).toLocaleDateString("tr-TR");

/** 05xx… / +90… → wa.me için 90xxxxxxxxxx; çözülemezse boş. */
function waNumara(tel: string): string {
  const d = String(tel || "").replace(/\D/g, "");
  if (!d) return "";
  if (d.startsWith("90") && d.length === 12) return d;
  if (d.startsWith("0") && d.length === 11) return "9" + d;
  if (d.length === 10) return "90" + d;
  return d.length >= 10 ? d : "";
}

interface Summary {
  orderCount: number;
  totalAmount: number;
  totalPaid: number;
  openingBalance: number;
  openingAsOf: string | null;
  balance: number;
  lastOrderAt: string | null;
}

export default function CustomerAccount({ id, finans = false }: { id: string; finans?: boolean }) {
  const router = useRouter();
  const [customer, setCustomer] = useState<Customer | null>(null);
  const [entries, setEntries] = useState<CariEntry[]>([]);
  const [movements, setMovements] = useState<Tahsilat[]>([]);
  const [summary, setSummary] = useState<Summary | null>(null);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [modalOpen, setModalOpen] = useState(false);
  const [editOpen, setEditOpen] = useState(false);
  const [msg, setMsg] = useState("");
  const [bolgeTanim, setBolgeTanim] = useState<Record<Bolge, BolgeInfo>>(() => varsayilanBolgeler());

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
          if (d.bolgeler) setBolgeTanim(d.bolgeler);
          setErr("");
        }
      })
      .catch(() => setErr("Sunucuya ulaşılamadı"))
      .finally(() => setLoading(false));
  }, [id]);

  useEffect(() => { load(); }, [load]);

  async function sil() {
    if (!customer) return;
    if (!confirm(`${customerTitle(customer)} kaydı silinsin mi? Sipariş kayıtları silinmez, yalnızca müşteri kartı kaldırılır.`)) return;
    const r = await fetch(`/api/musteriler?id=${encodeURIComponent(customer.id)}`, { method: "DELETE" });
    if (r.ok) router.push("/musteriler");
    else setErr("Silinemedi.");
  }

  if (loading) {
    return (
      <div className="stack" aria-busy="true">
        <div className="skeleton" style={{ height: 96, borderRadius: 16 }} />
        <div className="skeleton" style={{ height: 90, borderRadius: 16 }} />
        <div className="skeleton" style={{ height: 180, borderRadius: 16 }} />
      </div>
    );
  }
  if (err && !customer) {
    return (
      <div className="card">
        <div className="empty">
          <div className="empty-icon"><Icon name="alert" size={22} /></div>
          <strong>{err}</strong>
          <Link href="/musteriler" className="btn secondary small" style={{ marginTop: 10 }}>
            <Icon name="chevron-left" size={14} /> Müşteriler
          </Link>
        </div>
      </div>
    );
  }
  if (!customer) return null;

  const bakiye = summary?.balance ?? 0;
  const kisi = `${customer.firstName || ""} ${customer.lastName || ""}`.trim();
  const baslik = customer.company || kisi || customerTitle(customer);
  const adres = [customer.addr1, customer.addr2, [customer.district, customer.city].filter(Boolean).join(" / "), customer.postalCode].filter(Boolean).join(", ");
  const wa = waNumara(customer.phone);
  const bolge = musteriBolgesi(customer);
  const bolgeBilgi = bolgeTanim[bolge];

  return (
    <div>
      {/* ---- Başlık: avatar, ad, iletişim çipleri, hızlı işlemler ---- */}
      <div className="ck-head">
        <span className={`ck-avatar ${bolge}`} aria-hidden>{initials(baslik)}</span>
        <div className="ck-main">
          <span className="ck-kicker">Müşteri Kartı</span>
          <h1 className="ck-name">
            {baslik}
            <span className={`bolge ${bolge}`} title={customer.bolge ? "Bölge elle seçildi" : "Bölge şehirden türetildi"}>{bolgeBilgi.kisa}</span>
            {customer.iskontoPct ? <span className="badge brand">%{customer.iskontoPct} bayi iskontosu</span> : null}
          </h1>
          <div className="ck-person">
            {customer.company && kisi && <>Yetkili: {kisi} · </>}
            {bolgeBilgi.label} müşterisi · gönderici şube {customer.branch === "istanbul" ? "İstanbul" : "Ankara"}
          </div>
          <div className="ck-contacts">
            {customer.phone ? (
              <>
                <a className="ck-chip" href={`tel:${customer.phone.replace(/\s/g, "")}`}><Icon name="phone" size={14} /> {customer.phone}</a>
                {wa && (
                  <a className="ck-chip wa" href={`https://wa.me/${wa}`} target="_blank" rel="noopener noreferrer"><Icon name="message" size={14} /> WhatsApp</a>
                )}
              </>
            ) : (
              <span className="ck-chip plain"><Icon name="phone" size={14} /> Telefon girilmemiş</span>
            )}
            {customer.email && <a className="ck-chip" href={`mailto:${customer.email}`}><Icon name="mail" size={14} /> {customer.email}</a>}
          </div>
          <div className="ck-meta">
            <span><Icon name="map-pin" size={15} /> {adres || "Adres girilmemiş"}</span>
            {customer.note && <span><Icon name="info" size={15} /> {customer.note}</span>}
          </div>
        </div>
        <div className="ck-actions">
          <Link href={`/panel?musteri=${encodeURIComponent(customer.id)}`} className="btn">
            <Icon name="plus" size={16} /> Yeni Sipariş
          </Link>
          <Link href={`/etiket?id=${encodeURIComponent(customer.id)}`} className="btn secondary">
            <Icon name="tag" size={16} /> Kargo Etiketi
          </Link>
          <button type="button" className="btn secondary" onClick={() => setEditOpen(true)}>
            <Icon name="edit" size={16} /> Düzenle
          </button>
          {finans && (
            <button type="button" className="btn secondary" onClick={() => setModalOpen(true)}>
              <Icon name="wallet" size={16} /> Tahsilat Ekle
            </button>
          )}
          <Link href="/musteriler" className="btn ghost">
            <Icon name="chevron-left" size={16} /> Müşteriler
          </Link>
        </div>
      </div>

      {msg && <div className="notice ok">{msg}</div>}
      {err && <div className="notice err">{err}</div>}

      {/* ---- Özet kutuları ---- */}
      <div className="cari-cards">
        <div className="cari-card">
          <span>Sipariş Sayısı</span>
          <strong>{summary?.orderCount ?? 0}</strong>
        </div>
        <div className="cari-card">
          <span>Toplam Ciro</span>
          <strong>₺{fmt(summary?.totalAmount ?? 0)}</strong>
        </div>
        {finans ? (
          <>
            <div className="cari-card">
              <span>Tahsil Edilen</span>
              <strong style={{ color: "var(--success)" }}>₺{fmt(summary?.totalPaid ?? 0)}</strong>
            </div>
            <div className={`cari-card ${bakiye > 0 ? "borc" : ""}`}>
              <span>Kalan Bakiye</span>
              <strong style={{ color: bakiye > 0 ? "var(--error)" : "var(--success)" }}>₺{fmt(bakiye)}</strong>
            </div>
          </>
        ) : (
          <div className="cari-card">
            <span>Son Sipariş</span>
            <strong>{summary?.lastOrderAt ? tarih(summary.lastOrderAt) : "—"}</strong>
          </div>
        )}
      </div>

      {finans && summary && summary.openingBalance !== 0 && (
        <p className="notice info">
          Devir (açılış) bakiyesi: <strong>₺{fmt(summary.openingBalance)}</strong>
          {summary.openingAsOf && <> — {summary.openingAsOf} tarihi itibarıyla, Excel&apos;den aktarıldı.</>}
        </p>
      )}

      {/* ---- Mikro (resmi) cari: eşleştirme + canlı bakiye ---- */}
      <MikroCariKutusu customerId={customer.id} />

      {/* ---- Tahsilat hareketleri (finans modülü) ---- */}
      {finans && <div className="card" style={{ marginTop: 18 }}>
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
                  <th>Tarih</th><th>Yöntem</th><th>Şube</th><th className="num">Tutar</th><th>Sipariş</th><th>Kaydeden</th><th>Not</th>
                </tr>
              </thead>
              <tbody>
                {movements.map((t) => (
                  <tr key={t.id}>
                    <td style={{ whiteSpace: "nowrap" }}>{t.dateKey.split("-").reverse().join(".")}</td>
                    <td style={{ whiteSpace: "nowrap" }}>{TAHSILAT_YONTEM_LABELS[t.method] || t.method}</td>
                    <td style={{ fontSize: 12.5 }}>{t.branch === "istanbul" ? "İST" : "ANK"}</td>
                    <td className="num" style={{ color: "var(--success)", fontWeight: 600, whiteSpace: "nowrap" }}>
                      {t.currency === "TL" ? "₺" : t.currency === "USD" ? "$" : "€"}{fmt(t.amount)}
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
      </div>}

      {/* ---- Sipariş geçmişi ---- */}
      <div className="card" style={{ marginTop: 18 }}>
        <div className="card-head">
          <span className="card-head-icon"><Icon name="list" size={16} /></span>
          <div>
            <h2>Sipariş Geçmişi</h2>
            <span className="card-head-sub">{entries.length} sipariş · bu sistemden alınanlar</span>
          </div>
          <span className="spacer" />
          <Link href={`/panel?musteri=${encodeURIComponent(customer.id)}`} className="btn small secondary">
            <Icon name="plus" size={14} /> Sipariş
          </Link>
        </div>
        {entries.length === 0 ? (
          <div className="empty" style={{ padding: "18px 12px" }}>
            <strong>Bu müşteriye ait sipariş yok</strong>
            Sipariş alırken müşteriyi listeden seçerseniz kayıtlar buraya düşer.
          </div>
        ) : (
          <>
            <div className="table-wrap ck-table">
              <table>
                <thead>
                  <tr>
                    <th>Sipariş</th>
                    <th>Tür</th>
                    <th>Tarih</th>
                    <th>Durum</th>
                    <th className="num">Tutar</th>
                    {finans && <th className="num">Tahsilat</th>}
                    {finans && <th className="num">Bakiye</th>}
                    <th className="no-print"></th>
                  </tr>
                </thead>
                <tbody>
                  {entries.map((e) => (
                    <tr key={`${e.kind}-${e.orderId}`}>
                      <td style={{ fontWeight: 700, color: "var(--brand)", whiteSpace: "nowrap" }}>{e.orderId}</td>
                      <td><span className={`badge ${e.kind === "toptan" ? "var" : "az"}`}>{e.kind === "toptan" ? "Toptan" : "Perakende"}</span></td>
                      <td style={{ whiteSpace: "nowrap" }}>{tarih(e.createdAt)}</td>
                      <td style={{ fontSize: 12.5 }}>{e.status}</td>
                      <td className="num" style={{ whiteSpace: "nowrap", fontWeight: 600 }}>₺{fmt(e.total)}</td>
                      {finans && <td className="num" style={{ color: "var(--success)", whiteSpace: "nowrap" }}>₺{fmt(e.paid)}</td>}
                      {finans && (
                        <td className="num" style={{ fontWeight: 700, whiteSpace: "nowrap", color: e.balance > 0 ? "var(--error)" : "var(--success)" }}>₺{fmt(e.balance)}</td>
                      )}
                      <td className="no-print">
                        <Link className="btn small secondary" href={fisHref(e)}>
                          <Icon name="file-text" size={14} /> Fiş
                        </Link>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {/* Telefon: kart listesi */}
            <div className="ck-orders">
              {entries.map((e) => (
                <Link key={`m-${e.kind}-${e.orderId}`} href={fisHref(e)} className="ck-order">
                  <span className="ck-order-no">{e.orderId}</span>
                  <span className="ck-order-amt">₺{fmt(e.total)}</span>
                  <span className="ck-order-sub">
                    <span className={`badge ${e.kind === "toptan" ? "var" : "az"}`}>{e.kind === "toptan" ? "Toptan" : "Perakende"}</span>
                    <span>{tarih(e.createdAt)}</span>
                    <span>· {e.status}</span>
                    {finans && e.balance > 0 && <span style={{ color: "var(--error)", fontWeight: 700 }}>· Kalan ₺{fmt(e.balance)}</span>}
                  </span>
                </Link>
              ))}
            </div>
          </>
        )}
      </div>

      <p className="muted" style={{ fontSize: 12.5, marginTop: 14 }}>
        Kayıt {new Date(customer.createdAt).toLocaleDateString("tr-TR")} tarihinde açıldı
        {customer.updatedAt && <>, son güncelleme {new Date(customer.updatedAt).toLocaleDateString("tr-TR")}</>}.{" "}
        <button type="button" className="btn ghost xs" onClick={sil} style={{ color: "var(--error)" }}>
          <Icon name="x" size={12} /> Müşteri kartını sil
        </button>
      </p>

      {editOpen && (
        <CustomerForm
          initial={customer}
          onClose={() => setEditOpen(false)}
          onSaved={(c) => { setCustomer(c); setEditOpen(false); setMsg("Müşteri bilgileri güncellendi."); setTimeout(() => setMsg(""), 3000); }}
        />
      )}

      {finans && modalOpen && (
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

function fisHref(e: CariEntry): string {
  return e.kind === "toptan"
    ? `/panel/siparisler/detay?d=${e.dateKey}&id=${encodeURIComponent(e.orderId)}`
    : `/panel/perakende/siparisler/detay?d=${e.dateKey}&id=${encodeURIComponent(e.orderId)}`;
}
