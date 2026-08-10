"use client";

// Teknik malzeme seçici — 117 ürünlük açılır liste kaydırmakla zor
// bulunuyordu. Yazarak süzülür (Türkçe karakter/boşluk duyarsız), ok
// tuşlarıyla gezilir, Enter ile seçilir. Ürünün görseli tanımlıysa
// listede küçük önizleme çıkar.

import { useEffect, useMemo, useRef, useState } from "react";
import {
  TECHNICAL_PRODUCTS,
  getTechnicalProduct,
  type TechnicalProduct,
} from "@/data/technical";
import { eslesir } from "@/lib/search-norm";

const fiyatEtiketi = (t: TechnicalProduct) =>
  t.priceTL != null
    ? `₺${t.priceTL.toLocaleString("tr-TR")}`
    : t.priceEUR != null
      ? `€${t.priceEUR.toLocaleString("tr-TR")}`
      : "";

export default function TechnicalPicker({
  value,
  onPick,
}: {
  value: string; // seçili ürün kodu
  onPick: (t: TechnicalProduct) => void;
}) {
  const secili = getTechnicalProduct(value);
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const [aktif, setAktif] = useState(0);
  const boxRef = useRef<HTMLDivElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);

  // Dışarı tıklayınca kapan, seçili ürüne geri dön
  useEffect(() => {
    function onDoc(e: MouseEvent) {
      if (boxRef.current && !boxRef.current.contains(e.target as Node)) {
        setOpen(false);
        setQuery("");
      }
    }
    document.addEventListener("click", onDoc);
    return () => document.removeEventListener("click", onDoc);
  }, []);

  const sonuclar = useMemo(() => {
    const q = query.trim();
    if (!q) return TECHNICAL_PRODUCTS;
    return TECHNICAL_PRODUCTS.filter((t) => eslesir(q, t.name, t.code, t.category));
  }, [query]);

  // Aktif satır listeden taşmasın
  useEffect(() => {
    if (!open) return;
    const el = listRef.current?.querySelector<HTMLElement>(`[data-i="${aktif}"]`);
    el?.scrollIntoView({ block: "nearest" });
  }, [aktif, open]);

  function sec(t: TechnicalProduct) {
    onPick(t);
    setOpen(false);
    setQuery("");
  }

  function tus(e: React.KeyboardEvent) {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setOpen(true);
      setAktif((i) => Math.min(i + 1, sonuclar.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setAktif((i) => Math.max(i - 1, 0));
    } else if (e.key === "Enter") {
      if (open && sonuclar[aktif]) {
        e.preventDefault();
        sec(sonuclar[aktif]);
      }
    } else if (e.key === "Escape") {
      setOpen(false);
      setQuery("");
    }
  }

  // Kategori başlıkları: aynı kategorinin ilk ürününden önce yazılır
  let oncekiKategori = "";

  return (
    <div className="cp-wrap tp-wrap" ref={boxRef}>
      <input
        value={open ? query : secili?.name || ""}
        onChange={(e) => {
          setQuery(e.target.value);
          setAktif(0);
          setOpen(true);
        }}
        onFocus={() => {
          setOpen(true);
          setQuery("");
          setAktif(0);
        }}
        onKeyDown={tus}
        placeholder={secili ? secili.name : "Ürün ara: agraf, karton, tel…"}
        autoComplete="off"
      />
      {open && (
        <div className="cp-menu tp-menu" ref={listRef}>
          {sonuclar.length === 0 ? (
            <div className="cp-empty">
              “{query}” için ürün bulunamadı.
            </div>
          ) : (
            sonuclar.map((t, i) => {
              const yeniKategori = t.category !== oncekiKategori;
              oncekiKategori = t.category;
              return (
                <div key={t.code}>
                  {yeniKategori && <div className="tp-cat">{t.category}</div>}
                  <button
                    type="button"
                    data-i={i}
                    className={`cp-item tp-item ${i === aktif ? "aktif" : ""} ${
                      t.code === value ? "secili" : ""
                    }`}
                    onMouseEnter={() => setAktif(i)}
                    onClick={() => sec(t)}
                  >
                    {t.image && (
                      // eslint-disable-next-line @next/next/no-img-element
                      <img className="tp-img" src={t.image} alt="" loading="lazy" />
                    )}
                    <span className="tp-ad">{t.name}</span>
                    <span className="tp-fiyat">{fiyatEtiketi(t)}</span>
                  </button>
                </div>
              );
            })
          )}
        </div>
      )}
    </div>
  );
}
