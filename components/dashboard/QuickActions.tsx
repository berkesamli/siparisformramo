/* eslint-disable @next/next/no-img-element */
import Link from "next/link";
import Icon, { type IconName } from "@/components/shell/Icon";

export interface QuickTile {
  title: string;
  sub: string;
  href: string;
  icon: IconName;
  img?: string | null;
}

export default function QuickActions({ tiles, chips }: { tiles: QuickTile[]; chips?: { label: string; href: string; icon: IconName }[] }) {
  return (
    <div>
      <div className="quick-grid">
        {tiles.map((t) => (
          <Link key={t.href} href={t.href} className="quick-tile">
            <span className="quick-media">
              {t.img ? <img src={t.img} alt="" /> : <span className="quick-fallback"><Icon name={t.icon} size={34} /></span>}
              <span className="quick-shade" />
            </span>
            <span className="quick-txt">
              <span className="quick-icon"><Icon name={t.icon} size={18} /></span>
              <strong>{t.title}</strong>
              <span>{t.sub}</span>
            </span>
            <span className="quick-arrow" aria-hidden><Icon name="arrow-up-right" size={18} /></span>
          </Link>
        ))}
      </div>
      {chips && chips.length > 0 && (
        <div className="row" style={{ marginTop: 14 }}>
          {chips.map((c) => (
            <Link key={c.href} href={c.href} className="chip">
              <Icon name={c.icon} size={14} /> {c.label}
            </Link>
          ))}
        </div>
      )}
    </div>
  );
}
