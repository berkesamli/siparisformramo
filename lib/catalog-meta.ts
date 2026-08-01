// Katalog kartlarında/görüntüleyicide dosya adına göre özel başlık ve not.
// Yeni PDF eklerken buraya kayıt eklemek zorunlu değildir — eşleşme yoksa
// başlık dosya adından üretilir.

export interface CatalogMeta {
  title: string;
  note?: string;
}

export const CATALOG_META: Record<string, CatalogMeta> = {
  "cerceve-profil-katalogu": {
    title: "Çerçeve Profilleri Toptan Fiyat Listesi",
    note: "Fiyatlara KDV dahil değildir.",
  },
  "teknik-malzeme-katalogu": {
    title: "Teknik Malzeme Kataloğu",
    note: "Fiyatlar + KDV'dir.",
  },
};

export function catalogTitle(slug: string): CatalogMeta {
  const meta = CATALOG_META[slug];
  if (meta) return meta;
  return {
    title: slug.replace(/[-_]+/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()),
  };
}
