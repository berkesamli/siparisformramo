"use client";

import { useEffect, useRef, useState } from "react";
import Link from "next/link";
import Icon from "./Icon";
import { initials } from "./nav-config";
import type { ShellUser } from "./types";

export default function UserMenu({ user, onLogout }: { user: ShellUser; onLogout: () => void }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") setOpen(false); };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  const yetkiler: string[] = [];
  if (user.owner) yetkiler.push("Sahip");
  if (user.finance) yetkiler.push("Finans");
  if (user.maliyet) yetkiler.push("Maliyet");
  if (user.kur) yetkiler.push("Kur");

  return (
    <div className="menu-wrap" ref={ref}>
      <button type="button" className="tb-user" onClick={() => setOpen((o) => !o)} aria-haspopup="menu" aria-expanded={open} title={user.name}>
        <span className="avatar" aria-hidden>{initials(user.name)}</span>
        <span className="tb-user-text">
          <span className="tb-user-name">{user.name}</span>
          <span className="tb-user-role">{user.role === "staff" ? "Çalışan" : "Müşteri"}</span>
        </span>
        <Icon name="chevron-down" size={15} className="tb-user-text" />
      </button>
      {open && (
        <div className="menu" role="menu">
          <div className="menu-head">
            <strong>{user.name}</strong>
            <span>@{user.username} · {user.role === "staff" ? "Çalışan" : "Müşteri / Bayi"}{yetkiler.length ? " · " + yetkiler.join(", ") : ""}</span>
          </div>
          <Link href="/" className="menu-item" role="menuitem" onClick={() => setOpen(false)}>
            <Icon name="home" size={17} /> Gösterge Paneli
          </Link>
          {user.role === "staff" && (
            <Link href="/panel" className="menu-item" role="menuitem" onClick={() => setOpen(false)}>
              <Icon name="plus" size={17} /> Yeni Sipariş
            </Link>
          )}
          <Link href="/kataloglar" className="menu-item" role="menuitem" onClick={() => setOpen(false)}>
            <Icon name="book" size={17} /> Kataloglar
          </Link>
          <div className="menu-sep" />
          <button type="button" className="menu-item danger" role="menuitem" onClick={onLogout}>
            <Icon name="log-out" size={17} /> Çıkış Yap
          </button>
        </div>
      )}
    </div>
  );
}
