"use client";

// Giriş ekranındaki "Jarvis" küresi: binlerce noktadan oluşan, yavaşça dönen
// parçacık küresi. Yalnızca giriş ekranında ve yalnızca süs olarak kullanılır
// (bilgi taşımaz, etkileşimi yoktur).
// - Renk temanın --brand (altın) değişkeninden okunur; tema değişince yenilenir.
// - "Hareketi azalt" açıksa tek kare çizilir; sekme arka plana geçince durur.
// - Kare başına binlerce path yerine alfa/boyut kovalarında toplu fillRect
//   kullanılır: telefonda da ~30 fps'te akıcıdır. Küçük tuvalde nokta azaltılır.

import { useEffect, useRef } from "react";

interface Nokta { x: number; y: number; z: number; s: number; parla: boolean }

// Belirlenimci rastgele sayı (mulberry32): her açılışta aynı küre
function rnd32(seed: number) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function markaRengi(): [number, number, number] {
  const v = getComputedStyle(document.documentElement).getPropertyValue("--brand").trim();
  const m = /^#([0-9a-f]{6})$/i.exec(v);
  if (!m) return [224, 169, 79];
  const n = parseInt(m[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

// Küre yüzeyinde noktalar: %62'si kutuplarda yoğunlaşan dağılım (referans
// görseldeki üst/alt parlaklık), kalanı düzgün; ince bir kabuk + az dış saçılma.
function noktalar(adet: number): Nokta[] {
  const r = rnd32(7);
  const out: Nokta[] = [];
  for (let i = 0; i < adet; i++) {
    const phi = r() * Math.PI * 2;
    const theta = r() < 0.62 ? r() * Math.PI : Math.acos(1 - 2 * r());
    let rr = 0.985 + r() * 0.03;
    if (r() < 0.05) rr = 1 + r() * 0.07;
    out.push({
      x: rr * Math.sin(theta) * Math.cos(phi),
      y: rr * Math.cos(theta),
      z: rr * Math.sin(theta) * Math.sin(phi),
      s: 0.45 + r() * 1.05,
      parla: r() < 0.03,
    });
  }
  return out;
}

const KOVA = 8;                 // alfa kovası sayısı
const BOYUT = [0.9, 1.3, 1.8];  // nokta kenarı (px)

export default function ParticleSphere({ className, adet = 5200 }: { className?: string; adet?: number }) {
  const ref = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    const canvas = ref.current;
    const ctx = canvas?.getContext("2d");
    if (!canvas || !ctx) return;

    const pts = noktalar(adet);
    const azalt = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    let renk = markaRengi();
    let acik = document.documentElement.getAttribute("data-theme") === "light";
    let size = 0;
    let dpr = 1;
    let aci = 0.6;
    let son = 0;
    let raf = 0;
    const tilt = 0.3;
    const ct = Math.cos(tilt), st = Math.sin(tilt);

    const boyutla = () => {
      const w = canvas.clientWidth || 300;
      dpr = Math.min(2, window.devicePixelRatio || 1);
      size = w;
      canvas.width = Math.round(w * dpr);
      canvas.height = Math.round(w * dpr);
    };

    const ciz = (a: number) => {
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, size, size);
      const cx = size / 2, cy = size / 2, R = size * 0.42;
      const n = size < 320 ? Math.min(pts.length, 2800) : pts.length;

      // Zemin parıltısı
      const g = ctx.createRadialGradient(cx, cy, R * 0.3, cx, cy, R * 1.3);
      g.addColorStop(0, `rgba(${renk.join(",")},${acik ? 0.1 : 0.09})`);
      g.addColorStop(0.7, `rgba(${renk.join(",")},0.035)`);
      g.addColorStop(1, "rgba(0,0,0,0)");
      ctx.fillStyle = g;
      ctx.fillRect(0, 0, size, size);

      // Nokta rengi: koyu temada altın→beyaz, açık temada altın→koyu kahve karışımı
      const karisim = acik ? 0.25 : 0.55;
      const hedef = acik ? [60, 34, 10] : [255, 255, 255];
      const cr = Math.round(renk[0] + (hedef[0] - renk[0]) * karisim);
      const cg = Math.round(renk[1] + (hedef[1] - renk[1]) * karisim);
      const cb = Math.round(renk[2] + (hedef[2] - renk[2]) * karisim);

      const ca = Math.cos(a), sa = Math.sin(a);
      const kova: number[][] = Array.from({ length: KOVA * 3 }, () => []);
      const parlak: number[] = [];
      for (let i = 0; i < n; i++) {
        const p = pts[i];
        // Y ekseni etrafında dön, sonra X ekseninde hafif eğ; dik izdüşüm
        const x1 = p.x * ca + p.z * sa;
        const z1 = -p.x * sa + p.z * ca;
        const y2 = p.y * ct - z1 * st;
        const z2 = p.y * st + z1 * ct;
        const derin = (z2 + 1) / 2;                 // 0 arka … 1 ön
        const d = Math.min(1, Math.hypot(x1, y2));  // izdüşümde merkeze uzaklık
        const kenar = d * d * d;                     // kenarda parlayan halka
        const alfa = (0.12 + 0.88 * derin) * (0.28 + 0.72 * kenar);
        const px = cx + x1 * R, py = cy + y2 * R;
        if (p.parla && derin > 0.6) {
          parlak.push(px, py, alfa);
        } else {
          const ai = Math.min(KOVA - 1, Math.floor(alfa * KOVA));
          const si = p.s < 0.8 ? 0 : p.s < 1.2 ? 1 : 2;
          kova[ai * 3 + si].push(px, py);
        }
      }
      for (let ai = 0; ai < KOVA; ai++) {
        const alfa = ((ai + 0.5) / KOVA).toFixed(3);
        for (let si = 0; si < 3; si++) {
          const list = kova[ai * 3 + si];
          if (!list.length) continue;
          ctx.fillStyle = `rgba(${cr},${cg},${cb},${alfa})`;
          const s = BOYUT[si], h = s / 2;
          ctx.beginPath();
          for (let k = 0; k < list.length; k += 2) ctx.rect(list[k] - h, list[k + 1] - h, s, s);
          ctx.fill();
        }
      }
      // Birkaç büyük parıltı (yalnızca ön yüzde)
      for (let k = 0; k < parlak.length; k += 3) {
        ctx.fillStyle = `rgba(${cr},${cg},${cb},${Math.min(1, parlak[k + 2] + 0.2).toFixed(3)})`;
        ctx.beginPath();
        ctx.arc(parlak[k], parlak[k + 1], 1.6, 0, Math.PI * 2);
        ctx.fill();
      }
    };

    const dongu = (t: number) => {
      if (t - son >= 33) {                       // ~30 fps yeter, pil yormaz
        if (son) aci += (t - son) * 0.00011;     // bir tur ≈ 57 sn
        son = t;
        ciz(aci);
      }
      raf = requestAnimationFrame(dongu);
    };
    const basla = () => {
      cancelAnimationFrame(raf);
      son = 0;
      if (azalt) ciz(aci);
      else raf = requestAnimationFrame(dongu);
    };

    boyutla();
    basla();
    const ro = new ResizeObserver(() => { boyutla(); if (azalt) ciz(aci); });
    ro.observe(canvas);
    const mo = new MutationObserver(() => {
      renk = markaRengi();
      acik = document.documentElement.getAttribute("data-theme") === "light";
      if (azalt) ciz(aci);
    });
    mo.observe(document.documentElement, { attributes: true, attributeFilter: ["data-theme"] });
    const gorunum = () => { if (document.hidden) cancelAnimationFrame(raf); else basla(); };
    document.addEventListener("visibilitychange", gorunum);
    return () => {
      cancelAnimationFrame(raf);
      ro.disconnect();
      mo.disconnect();
      document.removeEventListener("visibilitychange", gorunum);
    };
  }, [adet]);

  return <canvas ref={ref} className={className} aria-hidden="true" />;
}
