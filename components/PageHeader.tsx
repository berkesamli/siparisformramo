// Tüm sayfalarda ortak başlık şeridi: ikon + başlık + açıklama + sağda eylemler.
// Sunucu bileşenlerinden de kullanılabilir (istemci kancası içermez).

import type { ReactNode } from "react";
import Icon, { type IconName } from "./shell/Icon";

export default function PageHeader({
  title,
  subtitle,
  icon,
  kicker,
  actions,
  children,
}: {
  title: ReactNode;
  subtitle?: ReactNode;
  icon?: IconName;
  kicker?: string;
  actions?: ReactNode;
  children?: ReactNode;
}) {
  return (
    <div className="page-head">
      {icon && (
        <span className="page-head-icon" aria-hidden>
          <Icon name={icon} size={22} />
        </span>
      )}
      <div className="page-head-main">
        {kicker && <span className="page-kicker">{kicker}</span>}
        <h1>{title}</h1>
        {subtitle && <p className="subtitle">{subtitle}</p>}
        {children}
      </div>
      {actions && <div className="page-head-actions no-print">{actions}</div>}
    </div>
  );
}
