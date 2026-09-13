"use client";

// Teknik malzeme seçici — 117 ürünlük açılır liste kaydırmakla zor
// bulunuyordu. Yazarak süzülür (Türkçe karakter/boşluk duyarsız), ok
// tuşlarıyla gezilir, Enter ile seçilir. Ürünün görseli tanımlıysa
// listede küçük önizleme çıkar.

import { useEffect, useMemo, useRef, useState } from "react";
// Ürün listesi prop ile gelir (fiyatlar istemci paketinde durmasın diye
// üst bileşen /api/katalog'dan çeker — bkz. lib/use-katalog).
import { teknikBul, type TechnicalProduct } from "@/lib/catalog-utils";
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
  products,
}: {
  value: string; // seçili ürün kodu
  onPick: (t: TechnicalProduct) => void;
  products: TechnicalProduct[];
}) {
  const secili = teknikBul(products, value);
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const [aktif, setAktif] = useState(0);
  const boxRef = useRef<HTMLDivElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  // Açılır liste hücreden geniş (300–420 px) ve sola hizalı; sağ kenardan
  // taşacaksa sola kaydırılır ki telefonda/kenar hücrede ekrandan çıkmasın
  // (kart overflow:visible olduğundan taşma sayfayı yatay kaydırırdı).
  const [kaydir, setKaydir] = useState(0);

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
    if (!q) return products;
    return products.filter((t) => eslesir(q, t.name, t.code, t.category));
  }, [query, products]);

  // Aktif satır listeden taşmasın
  useEffect(() => {
    if (!open) return;
    const el = listRef.current?.querySelector<HTMLElement>(`[data-i="${aktif}"]`);
    el?.scrollIntoView({ block: "nearest" });
  }, [aktif, open]);

  useEffect(() => {
    if (!open) return;
    const box = boxRef.current;
    const menu = listRef.current;
    if (!box || !menu) return;
    const r = box.getBoundingClientRect();
    const genislik = menu.offsetWidth;
    const ekran = document.documentElement.clientWidth;
    const kenar = 12;
    let s = 0;
    if (r.left + genislik > ekran - kenar) s = ekran - kenar - (r.left + genislik);
    if (r.left + s < kenar) s = kenar - r.left;
    setKaydir(Math.round(s));
  }, [open, sonuclar.length]);

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
        <div
          className="cp-menu tp-menu"
          ref={listRef}
          style={{ left: kaydir, maxHeight: "min(340px, 60dvh)" }}
        >
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
                    <span className="tp-ad" style={{ minWidth: 0 }}>
                      {t.name}
                    </span>
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
