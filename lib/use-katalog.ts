"use client";

// Toptan katalog verisini (çerçeve profilleri + teknik malzeme) /api/katalog
// ucundan bir kez çeker ve modül içi önbellekte tutar — aynı sayfadaki tüm
// bileşenler tek istekle paylaşır. Fiyat listesi bilinçli olarak istemci
// paketinde DEĞİLDİR (bkz. lib/catalog-utils).

import { useEffect, useState } from "react";
import type { FrameProfile, TechnicalProduct } from "@/lib/catalog-utils";

export interface Katalog {
  profiles: FrameProfile[];
  technical: TechnicalProduct[];
  /** Veri sunucudan geldi mi? (gelmeden datalist/fiyat önerileri boş kalır) */
  yuklendi: boolean;
}

let onbellek: Pick<Katalog, "profiles" | "technical"> | null = null;
let bekleyen: Promise<Pick<Katalog, "profiles" | "technical"> | null> | null = null;

async function katalogGetir(): Promise<Pick<Katalog, "profiles" | "technical"> | null> {
  try {
    const r = await fetch("/api/katalog");
    if (!r.ok) return null;
    const d = await r.json();
    if (!d?.ok) return null;
    return {
      profiles: Array.isArray(d.profiles) ? d.profiles : [],
      technical: Array.isArray(d.technical) ? d.technical : [],
    };
  } catch {
    return null;
  }
}

export function useKatalog(): Katalog {
  const [k, setK] = useState<Katalog>(() =>
    onbellek
      ? { ...onbellek, yuklendi: true }
      : { profiles: [], technical: [], yuklendi: false }
  );

  useEffect(() => {
    if (onbellek) return;
    let aktif = true;
    if (!bekleyen) bekleyen = katalogGetir();
    bekleyen.then((d) => {
      if (!d) {
        bekleyen = null; // bir sonraki bileşen tekrar denesin
        return;
      }
      onbellek = d;
      if (aktif) setK({ ...d, yuklendi: true });
    });
    return () => {
      aktif = false;
    };
  }, []);

  return k;
}
