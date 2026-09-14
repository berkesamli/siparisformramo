import fs from "fs";
import path from "path";
import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { catalogTitle } from "@/lib/catalog-meta";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

function listCatalogs(): {
  slug: string;
  name: string;
  note?: string;
  sizeMB: string;
}[] {
  const dir = path.join(process.cwd(), "public", "catalogs");
  try {
    return fs
      .readdirSync(dir)
      .filter((f) => f.toLowerCase().endsWith(".pdf"))
      .map((f) => {
        const stat = fs.statSync(path.join(dir, f));
        const rawSlug = f.replace(/\.pdf$/i, "");
        const meta = catalogTitle(rawSlug);
        return {
          slug: encodeURIComponent(rawSlug),
          name: meta.title,
          note: meta.note,
          sizeMB: (stat.size / (1024 * 1024)).toFixed(1),
        };
      });
  } catch {
    return [];
  }
}

export default async function CatalogsPage() {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/kataloglar");

  const catalogs = listCatalogs();

  return (
    <main className="container">
      <PageHeader
        title="Kataloglar"
        subtitle="Ürün ve teknik malzeme kataloglarımızı dergi formatında sayfa çevirerek inceleyebilirsiniz."
        icon="book"
      />

      {catalogs.length === 0 ? (
        <div className="card">
          <div className="notice info">
            Henüz katalog yüklenmedi. PDF dosyalarınızı projenin{" "}
            <code>public/catalogs/</code> klasörüne ekleyin (örn.{" "}
            <code>toptan-fiyat-listesi.pdf</code>,{" "}
            <code>teknik-malzeme-katalogu.pdf</code>) — bu sayfada otomatik
            olarak dergi görünümüyle listelenirler.
          </div>
        </div>
      ) : (
        <div className="grid cols-3">
          {catalogs.map((c) => (
            <div className="card" key={c.slug} style={{ display: "flex", flexDirection: "column" }}>
              <span
                className="card-head-icon"
                style={{ width: 44, height: 44, borderRadius: 12, marginBottom: 12 }}
                aria-hidden
              >
                <Icon name="book" size={22} />
              </span>
              <h2 style={{ marginTop: 0 }}>{c.name}</h2>
              {c.note && (
                <p className="muted" style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 6 }}>
                  {c.note}
                </p>
              )}
              <p className="muted" style={{ fontSize: 13, marginBottom: 14 }}>
                PDF · {c.sizeMB} MB
              </p>
              <Link href={`/kataloglar/${c.slug}`} className="btn" style={{ marginTop: "auto", alignSelf: "flex-start" }}>
                <Icon name="eye" size={16} /> Dergi Görünümünde Aç
              </Link>
            </div>
          ))}
        </div>
      )}
    </main>
  );
}
