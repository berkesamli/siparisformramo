"use client";

/* eslint-disable @next/next/no-img-element */

import { useState, Suspense } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import Icon from "@/components/shell/Icon";
import ThemeToggle from "@/components/shell/ThemeToggle";
import ParticleSphere from "@/components/ParticleSphere";

function LoginForm() {
  const router = useRouter();
  const params = useSearchParams();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
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
      const next = params.get("next") || "/";
      router.push(next);
      router.refresh();
    } catch {
      setError("Sunucuya ulaşılamadı.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="login">
      <div className="login-theme"><ThemeToggle /></div>
      {/* Süs: altın parçacık küresi (yalnızca giriş ekranında; bilgi taşımaz) */}
      <div className="login-stage" aria-hidden="true">
        <ParticleSphere className="login-sphere" />
        <span className="login-stage-txt">SİPARİŞ · STOK · KATALOG · FİNANS</span>
      </div>
      <div className="login-card card">
        <div className="login-brand">
          <span className="login-mark"><img src="/logo.png" alt="Olga Çerçeve" /></span>
          <span className="login-word">olga</span>
          <span className="login-sub">ÇERÇEVE</span>
        </div>
        <h1>Yönetim Paneli</h1>
        <p className="subtitle" style={{ textAlign: "center", marginBottom: 22 }}>
          Çalışan veya bayi hesabınızla giriş yapın.
        </p>
        <form onSubmit={submit}>
          <div className="field">
            <label htmlFor="kadi">Kullanıcı Adı</label>
            <div className="login-input">
              <Icon name="user" size={17} />
              <input
                id="kadi"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                autoComplete="username"
                autoCapitalize="none"
                autoCorrect="off"
                required
              />
            </div>
          </div>
          <div className="field">
            <label htmlFor="sifre">Şifre</label>
            <div className="login-input">
              <Icon name="shield" size={17} />
              <input
                id="sifre"
                type={showPw ? "text" : "password"}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                autoComplete="current-password"
                required
              />
              <button
                type="button"
                className="btn icon ghost small"
                onClick={() => setShowPw((s) => !s)}
                aria-label={showPw ? "Şifreyi gizle" : "Şifreyi göster"}
                title={showPw ? "Şifreyi gizle" : "Şifreyi göster"}
              >
                <Icon name="eye" size={16} />
              </button>
            </div>
          </div>
          {error && <div className="notice err">{error}</div>}
          <button className="btn block" style={{ marginTop: 6 }} disabled={loading}>
            {loading ? "Giriş yapılıyor…" : "Giriş Yap"}
            {!loading && <Icon name="arrow-up-right" size={17} />}
          </button>
        </form>
        <p className="login-foot">
          Sipariş hattı <strong>0850 305 75 45</strong> · olgacerceve.com
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
