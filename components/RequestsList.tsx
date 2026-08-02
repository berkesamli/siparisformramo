"use client";

// Bayi taleplerinin çalışan görünümü: onayla / reddet, siparişe aktar.

import { useCallback, useEffect, useState } from "react";
import { REQUEST_LABELS, type SavedRequest, type RequestStatus } from "@/lib/requests";

export default function RequestsList() {
  const [requests, setRequests] = useState<SavedRequest[]>([]);
  const [loading, setLoading] = useState(true);
  const [filter, setFilter] = useState<RequestStatus | "all">("bekliyor");
  const [blobOk, setBlobOk] = useState(true);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/talepler");
      const d = await res.json();
      if (res.ok) {
        setRequests(d.requests || []);
        setBlobOk(d.blob !== false);
      }
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  async function setStatus(r: SavedRequest, status: RequestStatus) {
    const prev = requests;
    setRequests(requests.map((x) => (x.id === r.id ? { ...x, status } : x)));
    const res = await fetch(
      `/api/talepler?d=${r.dateKey}&id=${encodeURIComponent(r.id)}`,
      {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ status }),
      }
    );
    if (!res.ok) setRequests(prev);
  }

  const visible =
    filter === "all" ? requests : requests.filter((r) => r.status === filter);
  const pendingCount = requests.filter((r) => r.status === "bekliyor").length;

  return (
    <div>
      <div className="card" style={{ padding: 14, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
        {(["bekliyor", "onaylandi", "reddedildi", "all"] as const).map((f) => (
          <button
            key={f}
            className={`btn small ${filter === f ? "" : "secondary"}`}
            onClick={() => setFilter(f)}
          >
            {f === "all" ? "Tümü" : REQUEST_LABELS[f]}
            {f === "bekliyor" && pendingCount > 0 && ` (${pendingCount})`}
          </button>
        ))}
        <span style={{ flex: 1 }} />
        <button className="btn small secondary" onClick={load}>↻ Yenile</button>
      </div>

      {!blobOk && (
        <div className="notice info">
          Kalıcı depolama yapılandırılmadığı için talepler okunamıyor.
        </div>
      )}

      {loading ? (
        <p style={{ color: "var(--muted)" }}>Yükleniyor...</p>
      ) : visible.length === 0 ? (
        <div className="card" style={{ textAlign: "center", color: "var(--muted)" }}>
          Bu filtrede talep yok.
        </div>
      ) : (
        visible.map((r) => (
          <div className="card" key={r.id} style={{ padding: 16, marginBottom: 12 }}>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
              <strong style={{ color: "var(--brand)" }}>{r.customer}</strong>
              {r.phone && <span style={{ fontSize: 13 }}>{r.phone}</span>}
              <span style={{ color: "var(--muted)", fontSize: 12.5 }}>
                {new Date(r.createdAt).toLocaleString("tr-TR", {
                  dateStyle: "short",
                  timeStyle: "short",
                })}{" "}
                · {r.username}
              </span>
              <span style={{ flex: 1 }} />
              <span
                className={`badge ${
                  r.status === "onaylandi" ? "var" : r.status === "reddedildi" ? "yok" : "az"
                }`}
              >
                {REQUEST_LABELS[r.status]}
              </span>
            </div>

            <table style={{ marginTop: 10 }}>
              <thead>
                <tr>
                  <th>Ürün Kodu</th>
                  <th>Miktar</th>
                  <th>Not</th>
                </tr>
              </thead>
              <tbody>
                {r.lines.map((l, i) => (
                  <tr key={i}>
                    <td style={{ fontWeight: 600 }}>{l.code}</td>
                    <td>{l.qty} {l.unit}</td>
                    <td style={{ color: "var(--text-2)" }}>{l.note || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            {r.note && (
              <p style={{ fontSize: 13, color: "var(--text-2)", marginTop: 8 }}>
                <strong>Not:</strong> {r.note}
              </p>
            )}

            <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
              {r.status === "bekliyor" && (
                <>
                  <button className="btn small" onClick={() => setStatus(r, "onaylandi")}>
                    ✓ Onayla
                  </button>
                  <button className="btn small danger" onClick={() => setStatus(r, "reddedildi")}>
                    ✕ Reddet
                  </button>
                </>
              )}
              <a
                className="btn small secondary"
                href={`/panel?talep=${encodeURIComponent(
                  r.lines.map((l) => `${l.code} ${l.qty} ${l.unit}`).join(" | ")
                )}&musteri=${encodeURIComponent(r.customer)}`}
              >
                ➕ Siparişe Geç
              </a>
              {r.handledBy && (
                <span style={{ fontSize: 12, color: "var(--muted)", alignSelf: "center" }}>
                  İşleyen: {r.handledBy}
                </span>
              )}
            </div>
          </div>
        ))
      )}
    </div>
  );
}
