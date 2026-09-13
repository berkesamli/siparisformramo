"use client";

import Icon from "@/components/shell/Icon";

// Beklenmeyen hata sınırı — kabuk yerinde kalır, içerik alanında açıklama + yeniden dene.
export default function ErrorPage({ error, reset }: { error: Error & { digest?: string }; reset: () => void }) {
  return (
    <main className="container" style={{ maxWidth: 640 }}>
      <div className="card" style={{ textAlign: "center", padding: "44px 24px" }}>
        <div className="empty-icon" style={{ width: 64, height: 64, marginBottom: 14, color: "var(--error)", background: "var(--error-soft)" }}>
          <Icon name="alert" size={28} />
        </div>
        <h1 style={{ fontSize: 24 }}>Bir şeyler ters gitti</h1>
        <p className="subtitle" style={{ margin: "6px auto 20px" }}>
          Sayfa yüklenirken hata oluştu. Tekrar deneyin; sorun sürerse sayfayı yenileyin.
        </p>
        {error?.digest && <p className="muted small" style={{ marginBottom: 16 }}>Hata kodu: {error.digest}</p>}
        <div className="row" style={{ justifyContent: "center" }}>
          <button type="button" className="btn" onClick={reset}>
            <Icon name="refresh" size={16} /> Tekrar Dene
          </button>
          <a href="/" className="btn secondary">
            <Icon name="home" size={16} /> Gösterge Paneli
          </a>
        </div>
      </div>
    </main>
  );
}
