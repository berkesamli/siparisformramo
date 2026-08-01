import fs from "fs";
import path from "path";
import Link from "next/link";

export const dynamic = "force-dynamic";

function listCatalogs(): { slug: string; name: string; sizeMB: string }[] {
  const dir = path.join(process.cwd(), "public", "catalogs");
  try {
    return fs
      .readdirSync(dir)
      .filter((f) => f.toLowerCase().endsWith(".pdf"))
      .map((f) => {
        const stat = fs.statSync(path.join(dir, f));
        return {
          slug: encodeURIComponent(f.replace(/\.pdf$/i, "")),
          name: f
            .replace(/\.pdf$/i, "")
            .replace(/[-_]+/g, " ")
            .replace(/\b\w/g, (c) => c.toUpperCase()),
          sizeMB: (stat.size / (1024 * 1024)).toFixed(1),
        };
      });
  } catch {
    return [];
  }
}

export default function CatalogsPage() {
  const catalogs = listCatalogs();

  return (
    <main className="container">
      <h1>Kataloglar</h1>
      <p className="subtitle">
        Ürün ve teknik malzeme kataloglarımızı dergi formatında sayfa çevirerek
        inceleyebilirsiniz.
      </p>

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
            <div className="card" key={c.slug}>
              <div style={{ fontSize: 40, marginBottom: 8 }}>📖</div>
              <h2 style={{ marginTop: 0 }}>{c.name}</h2>
              <p style={{ color: "var(--muted)", fontSize: 13, marginBottom: 14 }}>
                PDF · {c.sizeMB} MB
              </p>
              <Link href={`/kataloglar/${c.slug}`} className="btn">
                Dergi Görünümünde Aç
              </Link>
            </div>
          ))}
        </div>
      )}
    </main>
  );
}
