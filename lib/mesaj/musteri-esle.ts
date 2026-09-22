// Gelen mesajın karşı tarafını (telefon / e-posta) müşteri kartına bağlar.
// Önce toptan müşteri defteri (C…), sonra perakende defteri (P…) denenir.
// Eşleşme yoksa konuşma "kayıtsız" kalır; çalışan panelden elle bağlayabilir.
//
// Web sitesi / perakende müşterileri toptan defterde olmadığı için çoğu
// WhatsApp (0850) ve Instagram konuşması ya perakende karta bağlanır ya da
// kayıtsız kalır — bu beklenen durumdur.

import { listCustomers, customerTitle } from "@/lib/customers";
import { listRetailCustomers, phoneKey } from "@/lib/retail-customers";
import { memo } from "@/lib/server-cache";

export interface MusteriEslesme {
  musteriId: string;
  musteriTur: "toptan" | "perakende";
  ad: string;
}

const epostaAnahtar = (s: string) => String(s || "").trim().toLowerCase();

async function dizinler() {
  return memo("mesaj:musteri-dizin", 60_000, async () => {
    const [toptan, perakende] = await Promise.all([
      listCustomers().catch(() => []),
      listRetailCustomers().catch(() => []),
    ]);
    const telefon = new Map<string, MusteriEslesme>();
    const eposta = new Map<string, MusteriEslesme>();
    // Perakende önce yazılır, toptan üstüne yazar → aynı numara iki defterde varsa toptan kazanır.
    for (const c of perakende) {
      const e: MusteriEslesme = { musteriId: c.id, musteriTur: "perakende", ad: c.name };
      const t = phoneKey(c.phone);
      if (t.length >= 10) telefon.set(t, e);
      const m = epostaAnahtar(c.email);
      if (m.includes("@")) eposta.set(m, e);
    }
    for (const c of toptan) {
      const e: MusteriEslesme = { musteriId: c.id, musteriTur: "toptan", ad: customerTitle(c) };
      const t = phoneKey(c.phone);
      if (t.length >= 10) telefon.set(t, e);
      const m = epostaAnahtar(c.email);
      if (m.includes("@")) eposta.set(m, e);
    }
    return { telefon, eposta };
  });
}

export async function musteriEsle(k: { telefon?: string; eposta?: string }): Promise<MusteriEslesme | null> {
  const d = await dizinler();
  if (k.telefon) {
    const t = phoneKey(k.telefon);
    if (t.length >= 10 && d.telefon.has(t)) return d.telefon.get(t)!;
  }
  if (k.eposta) {
    const m = epostaAnahtar(k.eposta);
    if (d.eposta.has(m)) return d.eposta.get(m)!;
  }
  return null;
}
