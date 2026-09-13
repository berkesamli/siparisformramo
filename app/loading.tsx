// Rota geçişlerinde içerik alanında hafif iskelet — kabuk yerinde kalır.
export default function Loading() {
  return (
    <main className="container" aria-busy="true" aria-live="polite">
      <div className="page-head">
        <div className="page-head-main">
          <div className="skeleton" style={{ width: 240, height: 28, marginBottom: 10 }} />
          <div className="skeleton" style={{ width: 360, height: 14 }} />
        </div>
      </div>
      <div className="card" style={{ minHeight: 220 }}>
        <div className="skeleton" style={{ width: "40%", height: 16, marginBottom: 14 }} />
        <div className="skeleton" style={{ width: "100%", height: 12, marginBottom: 8 }} />
        <div className="skeleton" style={{ width: "92%", height: 12, marginBottom: 8 }} />
        <div className="skeleton" style={{ width: "78%", height: 12 }} />
      </div>
    </main>
  );
}
