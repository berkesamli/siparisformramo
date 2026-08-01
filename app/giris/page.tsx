"use client";

import { useState, Suspense } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import Logo from "@/components/Logo";

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
      const next =
        params.get("next") || (data.role === "staff" ? "/panel" : "/portal");
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
        <Logo size={1.2} />
      </div>
      <div className="card">
        <h1 style={{ textAlign: "center", fontSize: 21 }}>Giriş Yap</h1>
        <p className="subtitle" style={{ textAlign: "center" }}>
          Çalışan veya bayi hesabınızla giriş yapın
        </p>
        <form onSubmit={submit}>
          <div style={{ marginBottom: 14 }}>
            <label>Kullanıcı Adı</label>
            <input
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              autoComplete="username"
              required
            />
          </div>
          <div style={{ marginBottom: 18 }}>
            <label>Şifre</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              autoComplete="current-password"
              required
            />
          </div>
          {error && <div className="notice err">{error}</div>}
          <button className="btn" style={{ width: "100%", justifyContent: "center" }} disabled={loading}>
            {loading ? "Giriş yapılıyor…" : "Giriş Yap"}
          </button>
        </form>
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
