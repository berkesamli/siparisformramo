"use client";

// Müşterinin Mikro'daki canlı bakiyesi — tek satırlık rozet. Komuta satırında
// ve arama paletinde müşteri satırının yanında tembel yüklenir; aynı müşteri
// için istek 60 sn boyunca paylaşılır (Mikro dış sistem, cevabı geç gelebilir).
// Mikro ayarlı değilse hiçbir şey çizmez; bağlı değilse kısa not düşer.

import { useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Durum {
  ok: boolean;
  kurulu?: boolean;
  bagli?: boolean;
  ozet?: { bakiye: number; borc: number; alacak: number; sonHareket: string | null };
  hata?: string;
  error?: string;
}

const TTL = 60_000;
const bellek = new Map<string, { ts: number; p: Promise<Durum> }>();

function yukle(customerId: string): Promise<Durum> {
  const simdi = Date.now();
  const h = bellek.get(customerId);
  if (h && simdi - h.ts < TTL) return h.p;
  const p = fetch(`/api/mikro/cari?musteri=${encodeURIComponent(customerId)}`)
    .then((r) => r.json() as Promise<Durum>)
    .catch(() => ({ ok: false, kurulu: true, hata: "Sunucuya ulaşılamadı." }));
  bellek.set(customerId, { ts: simdi, p });
  return p;
}

const fmt = (n: number) => "₺" + Math.abs(Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 });

export default function BakiyeChip({ customerId }: { customerId: string }) {
  const [d, setD] = useState<Durum | null>(null);

  useEffect(() => {
    let canli = true;
    setD(null);
    yukle(customerId).then((r) => { if (canli) setD(r); });
    return () => { canli = false; };
  }, [customerId]);

  if (!d) return <span className="bakiye yukleniyor skeleton" aria-label="Bakiye yükleniyor" />;
  if (d.kurulu === false) return null;
  if (d.bagli === false) {
    return (
      <span className="bakiye yok" title="Müşteri kartından Mikro cari kartıyla eşleştirin">
        <Icon name="wallet" size={12} /> Mikro&apos;ya bağlı değil
      </span>
    );
  }
  if (!d.ok || !d.ozet) {
    return <span className="bakiye hata" title={d.hata || d.error || ""}><Icon name="wallet" size={12} /> Bakiye alınamadı</span>;
  }
  const b = d.ozet.bakiye;
  const sinif = b > 0 ? "borclu" : b < 0 ? "alacakli" : "sifir";
  const yazi = b > 0 ? `${fmt(b)} borçlu` : b < 0 ? `${fmt(b)} alacaklı` : "Bakiye yok";
  return (
    <span className={`bakiye ${sinif}`} title={d.ozet.sonHareket ? `Son hareket ${d.ozet.sonHareket}` : "Mikro bakiyesi"}>
      <Icon name="wallet" size={12} /> {yazi}
    </span>
  );
}
