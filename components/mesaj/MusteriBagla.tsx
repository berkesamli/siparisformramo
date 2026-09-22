"use client";

// Konuşmayı toptan (C…) ya da perakende (P…) müşteri kartına bağlama modalı.

import { useEffect, useMemo, useState } from "react";
import Icon from "@/components/shell/Icon";
import { eslesir } from "@/lib/search-norm";

interface Toptan { id: string; firstName: string; lastName: string; company: string; phone: string; email: string; city: string }
interface Perakende { id: string; name: string; phone: string; email: string }

export default function MusteriBagla({ mevcut, ipucu, onClose, onSelect }: {
  mevcut: { id: string; tur: "toptan" | "perakende"; ad: string } | null;
  ipucu: string;
  onClose: () => void;
  onSelect: (id: string | null, tur: "toptan" | "perakende" | null) => Promise<void>;
}) {
  const [q, setQ] = useState("");
  const [toptan, setToptan] = useState<Toptan[]>([]);
  const [perakende, setPerakende] = useState<Perakende[]>([]);
  const [yukleniyor, setYukleniyor] = useState(true);
  const [kaydediliyor, setKaydediliyor] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const [a, b] = await Promise.all([
          fetch("/api/musteriler", { cache: "no-store" }).then((r) => r.json()).catch(() => ({})),
          fetch("/api/perakende/musteriler", { cache: "no-store" }).then((r) => r.json()).catch(() => ({})),
        ]);
        setToptan(Array.isArray(a?.customers) ? a.customers : []);
        setPerakende(Array.isArray(b?.customers) ? b.customers : []);
      } finally { setYukleniyor(false); }
    })();
  }, []);

  // İlk açılışta ipucu (numara/e-posta/ad) ile ara
  useEffect(() => { if (ipucu && !q) setQ(ipucu.replace(/^\+?90/, "").slice(0, 40)); /* eslint-disable-line react-hooks/exhaustive-deps */ }, [ipucu]);

  const t = useMemo(() => {
    const s = q.trim();
    const l = s ? toptan.filter((c) => eslesir(s, c.company, c.firstName, c.lastName, c.phone, c.email, c.city)) : toptan;
    return l.slice(0, 30);
  }, [q, toptan]);
  const p = useMemo(() => {
    const s = q.trim();
    const l = s ? perakende.filter((c) => eslesir(s, c.name, c.phone, c.email)) : perakende;
    return l.slice(0, 30);
  }, [q, perakende]);

  async function sec(id: string | null, tur: "toptan" | "perakende" | null) {
    setKaydediliyor(true);
    try { await onSelect(id, tur); } finally { setKaydediliyor(false); }
  }

  return (
    <div className="backdrop ib-modal-wrap" onClick={onClose} role="dialog" aria-modal="true" aria-label="Müşteri bağla">
      <div className="ib-modal" onClick={(e) => e.stopPropagation()}>
        <div className="ib-modal-head">
          <strong>Müşteri kartına bağla</strong>
          <button type="button" className="btn ghost icon small" onClick={onClose} aria-label="Kapat"><Icon name="x" size={16} /></button>
        </div>
        <div className="ib-search" style={{ margin: "0 0 10px" }}>
          <Icon name="search" size={16} />
          <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Firma, ad, telefon, e-posta…" aria-label="Müşteri ara" />
        </div>
        {mevcut && (
          <div className="notice" style={{ margin: "0 0 10px", display: "flex", alignItems: "center", gap: 8 }}>
            <span>Şu an bağlı: <strong>{mevcut.ad}</strong> ({mevcut.tur === "toptan" ? "bayi" : "perakende"})</span>
            <span className="spacer" />
            <button type="button" className="btn ghost xs" disabled={kaydediliyor} onClick={() => void sec(null, null)}>Bağlantıyı kaldır</button>
          </div>
        )}
        <div className={`ib-modal-body ${kaydediliyor ? "loading-dim" : ""}`}>
          {yukleniyor && <div className="empty">Müşteriler yükleniyor…</div>}
          {!yukleniyor && !t.length && !p.length && <div className="empty">Eşleşen müşteri yok. Müşteriyi önce <a href="/musteriler">Müşteriler</a> ya da <a href="/panel/perakende/musteriler">Perakende Müşteriler</a> sayfasından kaydedin.</div>}
          {t.length > 0 && <div className="ib-modal-grp">Bayiler (toptan)</div>}
          {t.map((c) => (
            <button key={c.id} type="button" className="ib-modal-row" onClick={() => void sec(c.id, "toptan")}>
              <span className="badge brand"><Icon name="briefcase" size={11} /></span>
              <span><strong>{c.company || `${c.firstName} ${c.lastName}`.trim()}</strong><small>{[c.firstName && c.company ? `${c.firstName} ${c.lastName}`.trim() : "", c.city, c.phone, c.email].filter(Boolean).join(" · ")}</small></span>
            </button>
          ))}
          {p.length > 0 && <div className="ib-modal-grp">Perakende müşteriler</div>}
          {p.map((c) => (
            <button key={c.id} type="button" className="ib-modal-row" onClick={() => void sec(c.id, "perakende")}>
              <span className="badge info"><Icon name="user" size={11} /></span>
              <span><strong>{c.name}</strong><small>{[c.phone, c.email].filter(Boolean).join(" · ")}</small></span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}
