import Link from "next/link";
import FlipBook from "@/components/FlipBook";
import { catalogTitle } from "@/lib/catalog-meta";

export const dynamic = "force-dynamic";

export default function CatalogViewerPage({
  params,
}: {
  params: { slug: string };
}) {
  const slug = decodeURIComponent(params.slug);
  const pdfUrl = `/catalogs/${encodeURIComponent(slug)}.pdf`;
  const meta = catalogTitle(slug);

  return (
    <main className="container" style={{ maxWidth: 1400 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 12, flexWrap: "wrap" }}>
        <Link href="/kataloglar" className="btn small secondary">
          ← Kataloglar
        </Link>
        <h1 style={{ fontSize: 20, margin: 0 }}>{meta.title}</h1>
        {meta.note && (
          <span style={{ color: "var(--brand)", fontSize: 13, fontWeight: 600 }}>
            {meta.note}
          </span>
        )}
      </div>
      <FlipBook pdfUrl={pdfUrl} />
    </main>
  );
}
