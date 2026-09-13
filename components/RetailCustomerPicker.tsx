"use client";

// Perakende müşteri seçici — PERAKENDE defterinden (retail-customers/)
// arar; etiket/toptan defterine hiç bakmaz. Yazarken süzülür, seçilince
// telefon/e-posta/adres forma dolar. Kayıtlı değilse yazılan ad kullanılır;
// sipariş kaydedilince müşteri deftere kendiliğinden işlenir.

import { useEffect, useMemo, useRef, useState } from "react";
import type { RetailCustomer } from "@/lib/retail-customers";
import { eslesir } from "@/lib/search-norm";

export default function RetailCustomerPicker({
  value,
  onChange,
  onPick,
  placeholder = "Müşteri adı",
}: {
  value: string;
  onChange: (v: string) => void;
  onPick?: (c: RetailCustomer) => void;
  placeholder?: string;
}) {
  const [customers, setCustomers] = useState<RetailCustomer[]>([]);
  const [open, setOpen] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const boxRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    fetch("/api/perakende/musteriler")
      .then((r) => (r.ok ? r.json() : null))
      .then((d) => {
        if (d?.customers) setCustomers(d.customers);
        setLoaded(true);
      })
      .catch(() => setLoaded(true));
  }, []);

  useEffect(() => {
    function onDoc(e: MouseEvent) {
      if (boxRef.current && !boxRef.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("click", onDoc);
    return () => document.removeEventListener("click", onDoc);
  }, []);

  const matches = useMemo(() => {
    const q = value.trim();
    const list = q
      ? customers.filter((c) => eslesir(q, c.name, c.phone, c.email))
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
            <div className="cp-empty">
              Eşleşen perakende müşterisi yok — yazdığınız ad kullanılır,
              sipariş kaydedilince deftere eklenir.
            </div>
          ) : (
            matches.map((c) => (
              <button
                key={c.id}
                type="button"
                className="cp-item"
                onClick={() => {
                  onChange(c.name);
                  onPick?.(c);
                  setOpen(false);
                }}
              >
                <span className="cp-name">{c.name}</span>
                <span className="cp-meta">
                  {[c.phone, c.email].filter(Boolean).join(" · ")}
                </span>
              </button>
            ))
          )}
          <a className="cp-foot" href="/panel/perakende/musteriler">
            + Perakende müşteri defterini aç
          </a>
        </div>
      )}
    </div>
  );
}
