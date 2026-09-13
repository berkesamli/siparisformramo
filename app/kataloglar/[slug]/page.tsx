import Link from "next/link";
import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import FlipBook from "@/components/FlipBook";
import { catalogTitle } from "@/lib/catalog-meta";
import PageHeader from "@/components/PageHeader";
import Icon from "@/components/shell/Icon";

export const dynamic = "force-dynamic";

export default async function CatalogViewerPage({
  params,
}: {
  params: { slug: string };
}) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/kataloglar");

  const slug = decodeURIComponent(params.slug);
  const pdfUrl = `/catalogs/${encodeURIComponent(slug)}.pdf`;
  const meta = catalogTitle(slug);

  return (
    <main className="container" style={{ maxWidth: 1400 }}>
      <PageHeader
        title={meta.title}
        subtitle={meta.note}
        icon="book"
        actions={
          <Link href="/kataloglar" className="btn secondary">
            <Icon name="chevron-left" size={16} /> Kataloglar
          </Link>
        }
      />
      <FlipBook pdfUrl={pdfUrl} />
    </main>
  );
}
