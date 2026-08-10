"use client";

// Kayıtlı müşterilerden seçim yapan otomatik tamamlama kutusu.
// Sipariş formlarında müşteri adı alanının yanında kullanılır; yazarken
// kayıtlı müşteriler süzülür, seçilince ad (ve varsa telefon/adres) döner.

import { useEffect, useMemo, useRef, useState } from "react";
import { customerTitle, normalizeCity, type Customer } from "@/lib/customers";
import { eslesir } from "@/lib/search-norm";

export default function CustomerPicker({
  value,
  onChange,
  onPick,
  placeholder = "Müşteri / Firma adı",
}: {
  value: string;
  onChange: (v: string) => void;
  onPick?: (c: Customer) => void;
  placeholder?: string;
}) {
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [open, setOpen] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const boxRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    fetch("/api/musteriler")
      .then((r) => (r.ok ? r.json() : null))
      .then((d) => {
        if (d?.customers) setCustomers(d.customers);
        setLoaded(true);
      })
      .catch(() => setLoaded(true));
  }, []);

  // Dışarı tıklayınca kapat
  useEffect(() => {
    function onDoc(e: MouseEvent) {
      if (boxRef.current && !boxRef.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("click", onDoc);
    return () => document.removeEventListener("click", onDoc);
  }, []);

  const matches = useMemo(() => {
    // eslesir(): Türkçe karakter ve büyük/küçük harf duyarsız arama.
    // Düz toLowerCase() yetmiyor — "YILMAZ".toLowerCase() "yilmaz" verir,
    // kullanıcı "yılmaz" yazınca eşleşmezdi.
    const q = value.trim();
    const list = q
      ? customers.filter((c) =>
          eslesir(q, customerTitle(c), c.company, c.firstName, c.lastName, c.phone, c.city)
        )
      : customers;
    return list.slice(0, 8);
  }, [customers, value]);

  return (
    <div className="cp-wrap" ref={boxRef}>
      <input
        value={value}
        onChange={(e) => {
          onChange(e.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
        placeholder={placeholder}
        autoComplete="off"
      />
      {open && loaded && customers.length > 0 && (
        <div className="cp-menu">
          {matches.length === 0 ? (
            <div className="cp-empty">Eşleşen kayıtlı müşteri yok — yazdığınız ad kullanılır.</div>
          ) : (
            matches.map((c) => (
              <button
                key={c.id}
                type="button"
                className="cp-item"
                onClick={() => {
                  onChange(customerTitle(c));
                  onPick?.(c);
                  setOpen(false);
                }}
              >
                <span className="cp-name">{customerTitle(c)}</span>
                <span className="cp-meta">
                  {[c.phone, [c.district, c.city].filter(Boolean).join(" / ")]
                    .filter(Boolean)
                    .join(" · ")}
                </span>
                {c.city && (
                  <span className={`cp-city ${normalizeCity(c.city)}`}>{c.city}</span>
                )}
              </button>
            ))
          )}
          <a className="cp-foot" href="/etiket">
            + Müşteri defterini aç
          </a>
        </div>
      )}
    </div>
  );
}
