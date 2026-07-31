import Link from "next/link";
import FlipBook from "@/components/FlipBook";

export const dynamic = "force-dynamic";

export default function CatalogViewerPage({
  params,
}: {
  params: { slug: string };
}) {
  const slug = decodeURIComponent(params.slug);
  const pdfUrl = `/catalogs/${encodeURIComponent(slug)}.pdf`;
  const title = slug.replace(/[-_]+/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());

  return (
    <main className="container" style={{ maxWidth: 1400 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 12 }}>
        <Link href="/kataloglar" className="btn small secondary">
          ← Kataloglar
        </Link>
        <h1 style={{ fontSize: 20, margin: 0 }}>{title}</h1>
      </div>
      <FlipBook pdfUrl={pdfUrl} />
    </main>
  );
}
