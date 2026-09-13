"use client";

import Link from "next/link";
import Icon from "@/components/shell/Icon";
import { STATUS_LABELS, PAYMENT_LABELS, type OrderIndexEntry } from "@/lib/orders";
import type { RetailIndexEntry } from "@/lib/retail-orders";

const tl = (n: number) => "₺" + (Number(n) || 0).toLocaleString("tr-TR", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
const saat = (iso: string) => new Date(iso).toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit", timeZone: "Europe/Istanbul" });
const gun = (k: string) => { const [, m, d] = k.split("-"); return `${Number(d)}.${m}`; };

const TOPTAN_BADGE: Record<string, string> = { olusturuldu: "warn", hazirlaniyor: "info", tamamlandi: "ok", iptal: "err" };
const PERAKENDE_BADGE: Record<string, string> = { Beklemede: "warn", "Hazırlanıyor": "info", "Hazır": "ok", "Teslim Edildi": "", "İptal": "err" };

export function RecentWholesale({ orders, loading, blob }: { orders: OrderIndexEntry[]; loading?: boolean; blob: boolean }) {
  if (!loading && orders.length === 0) {
    return (
      <div className="empty">
        <div className="empty-icon"><Icon name="inbox" size={22} /></div>
        <strong>Henüz sipariş yok</strong>
        {blob ? "Bu ay girilen toptan siparişler burada listelenir." : "Kalıcı depolama bağlandığında siparişler burada görünür."}
      </div>
    );
  }
  return (
    <ul className={`recent ${loading ? "loading-dim" : ""}`}>
      {orders.map((o) => (
        <li key={o.orderId}>
          <Link href={`/panel/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`} className="recent-row">
            <span className="recent-avatar">{(o.customer || "?").slice(0, 1).toLocaleUpperCase("tr-TR")}</span>
            <span className="recent-main">
              <span className="recent-title">{o.customer || "—"}</span>
              <span className="recent-sub">{o.orderId} · {gun(o.dateKey)} {saat(o.createdAt)} · {o.employee}{o.payment ? " · " + PAYMENT_LABELS[o.payment] : ""}</span>
            </span>
            <span className="recent-amt num">{tl(o.net)}</span>
            <span className={`badge ${TOPTAN_BADGE[o.status] || ""}`}>{STATUS_LABELS[o.status]}</span>
          </Link>
        </li>
      ))}
    </ul>
  );
}

export function RecentRetail({ orders, loading, blob }: { orders: RetailIndexEntry[]; loading?: boolean; blob: boolean }) {
  if (!loading && orders.length === 0) {
    return (
      <div className="empty">
        <div className="empty-icon"><Icon name="frame" size={22} /></div>
        <strong>Perakende sipariş yok</strong>
        {blob ? "Online çerçeve siparişleri burada listelenir." : "Kalıcı depolama bağlandığında görünür."}
      </div>
    );
  }
  return (
    <ul className={`recent ${loading ? "loading-dim" : ""}`}>
      {orders.map((o) => (
        <li key={o.orderId}>
          <Link href={`/panel/perakende/siparisler/detay?d=${o.dateKey}&id=${encodeURIComponent(o.orderId)}`} className="recent-row">
            <span className="recent-avatar retail">{(o.customerName || "?").slice(0, 1).toLocaleUpperCase("tr-TR")}</span>
            <span className="recent-main">
              <span className="recent-title">{o.customerName || "—"}</span>
              <span className="recent-sub">{o.orderId} · {gun(o.dateKey)} · {o.adet} kalem{o.deliveryDate ? ` · teslim ${o.deliveryDate}` : ""}</span>
            </span>
            <span className="recent-amt num">{tl(o.total)}</span>
            <span className={`badge ${PERAKENDE_BADGE[o.status] || ""}`}>{o.status}</span>
          </Link>
        </li>
      ))}
    </ul>
  );
}
