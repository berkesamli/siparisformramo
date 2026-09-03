"use client";

/* eslint-disable @next/next/no-img-element */
import { useState, Suspense } from "react";
import { useRouter, useSearchParams } from "next/navigation";

function LoginForm() {
  const router = useRouter();
  const params = useSearchParams();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const res = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        setError(data.error || "Giriş başarısız.");
        return;
      }
      const next = params.get("next") || (data.kind === "admin" ? "/yonetim" : "/panel");
      router.push(next);
      router.refresh();
    } catch {
      setError("Sunucuya ulaşılamadı.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="container" style={{ maxWidth: 440 }}>
      <div style={{ textAlign: "center", marginTop: 50, marginBottom: 26 }}>
        <img src="/logo.png" alt="Olga Çerçeve" style={{ height: 62, width: "auto" }} />
        <div style={{ fontSize: 12, fontWeight: 600, letterSpacing: "0.5em", marginLeft: "0.5em", color: "var(--brand)", marginTop: 8 }}>
          BAYİ PANELİ
        </div>
      </div>
      <div className="card">
        <h1 style={{ textAlign: "center", fontSize: 21 }}>Bayi Girişi</h1>
        <p className="subtitle" style={{ textAlign: "center" }}>
          Olga Çerçeve tarafından verilen bayi hesabınızla giriş yapın
        </p>
        <form onSubmit={submit}>
          <div style={{ marginBottom: 14 }}>
            <label>Kullanıcı Adı</label>
            <input value={username} onChange={(e) => setUsername(e.target.value)} autoComplete="username" required />
          </div>
          <div style={{ marginBottom: 18 }}>
            <label>Şifre</label>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="current-password" required />
          </div>
          {error && <div className="notice err">{error}</div>}
          <button className="btn" style={{ width: "100%", justifyContent: "center" }} disabled={loading}>
            {loading ? "Giriş yapılıyor…" : "Giriş Yap"}
          </button>
        </form>
        <p style={{ fontSize: 12.5, color: "var(--muted)", textAlign: "center", marginTop: 16 }}>
          Bayi hesabı için Olga Çerçeve: 0850 305 75 45
        </p>
      </div>
    </main>
  );
}

export default function LoginPage() {
  return (
    <Suspense>
      <LoginForm />
    </Suspense>
  );
}
