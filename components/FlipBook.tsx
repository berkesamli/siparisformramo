"use client";

// PDF'i dergi gibi çift sayfa gösteren görüntüleyici (pdfjs-dist ile).
// Ok tuşları ve butonlarla sayfa çevrilir; çevirme animasyonu CSS ile verilir.

import { useCallback, useEffect, useRef, useState } from "react";
import Icon from "@/components/shell/Icon";

type PdfDoc = {
  numPages: number;
  getPage: (n: number) => Promise<any>;
};

// Klavye kısayolları bir form alanında yazarken (üst çubuk araması vb.)
// sayfa çevirmesin.
function formAlanindaMi(t: EventTarget | null): boolean {
  const el = t as HTMLElement | null;
  if (!el || !el.tagName) return false;
  const tag = el.tagName;
  return (
    tag === "INPUT" ||
    tag === "TEXTAREA" ||
    tag === "SELECT" ||
    el.isContentEditable === true
  );
}

export default function FlipBook({ pdfUrl }: { pdfUrl: string }) {
  const [doc, setDoc] = useState<PdfDoc | null>(null);
  const [error, setError] = useState("");
  const [spread, setSpread] = useState(0); // 0 = kapak
  const [animKey, setAnimKey] = useState(0);
  const [direction, setDirection] = useState<"next" | "prev">("next");
  // İlk sayfa çizilene kadar sayfa kutuları iskelet (skeleton) olarak gösterilir.
  const [painted, setPainted] = useState(false);
  const leftRef = useRef<HTMLCanvasElement>(null);
  const rightRef = useRef<HTMLCanvasElement>(null);
  const renderTask = useRef(0);

  useEffect(() => {
    let cancelled = false;
    setPainted(false);
    (async () => {
      try {
        const pdfjs = await import("pdfjs-dist");
        pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
        const loaded = await pdfjs.getDocument(pdfUrl).promise;
        if (!cancelled) setDoc(loaded as unknown as PdfDoc);
      } catch (e) {
        console.error(e);
        if (!cancelled)
          setError(
            "PDF yüklenemedi. Dosyanın public/catalogs klasöründe olduğundan emin olun."
          );
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [pdfUrl]);

  // Yayılım (spread) → sayfa numaraları: 0 → [_,1], 1 → [2,3], 2 → [4,5] ...
  const pagesOf = useCallback(
    (s: number): [number | null, number | null] => {
      if (!doc) return [null, null];
      if (s === 0) return [null, 1];
      const left = s * 2;
      const right = left + 1;
      return [
        left <= doc.numPages ? left : null,
        right <= doc.numPages ? right : null,
      ];
    },
    [doc]
  );

  const maxSpread = doc ? Math.ceil(doc.numPages / 2) : 0;

  const renderPage = useCallback(
    async (pageNum: number | null, canvas: HTMLCanvasElement | null, task: number) => {
      if (!canvas) return;
      const ctx = canvas.getContext("2d");
      if (!ctx) return;
      if (!pageNum || !doc) {
        // Boş taraf (ör. kapağın solu): görünmez ama yayılımın genişliğini
        // korur. Rengi temadan alınır — canvas CSS değişkeni okuyamaz, bu
        // yüzden hesaplanmış stilden çözülür.
        canvas.width = 600;
        canvas.height = 850;
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        const zemin =
          typeof window !== "undefined"
            ? getComputedStyle(document.documentElement)
                .getPropertyValue("--surface-solid")
                .trim()
            : "";
        if (zemin) {
          ctx.fillStyle = zemin;
          ctx.fillRect(0, 0, canvas.width, canvas.height);
        }
        return;
      }
      const page = await doc.getPage(pageNum);
      if (task !== renderTask.current) return;
      const viewport = page.getViewport({ scale: 1.6 });
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: ctx, canvas, viewport }).promise;
    },
    [doc]
  );

  useEffect(() => {
    if (!doc) return;
    const task = ++renderTask.current;
    const [l, r] = pagesOf(spread);
    Promise.all([
      renderPage(l, leftRef.current, task),
      renderPage(r, rightRef.current, task),
    ]).then(
      () => {
        // Yalnızca güncel görev iskeleti kaldırır; geride kalan görevler yok sayılır.
        if (task === renderTask.current) setPainted(true);
      },
      (e) => console.error(e)
    );
  }, [doc, spread, pagesOf, renderPage]);

  const go = useCallback(
    (delta: number) => {
      setSpread((s) => {
        const next = Math.min(Math.max(0, s + delta), Math.max(0, maxSpread));
        if (next !== s) {
          setDirection(delta > 0 ? "next" : "prev");
          setAnimKey((k) => k + 1);
        }
        return next;
      });
    },
    [maxSpread]
  );

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (formAlanindaMi(e.target)) return;
      if (e.key === "ArrowRight") go(1);
      if (e.key === "ArrowLeft") go(-1);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [go]);

  const [l, r] = pagesOf(spread);

  if (error) return <div className="notice err">{error}</div>;

  return (
    <div className="card" style={{ textAlign: "center" }}>
      <style>{`
        .flip-stage {
          perspective: 2200px;
          display: flex;
          justify-content: center;
          align-items: flex-start;
          gap: 2px;
          overflow-x: auto;
          padding: 10px 0;
        }
        /* Sayfa genişliği kapsayıcıya göre: iki sayfa sahneye sığacak şekilde
           küçülür (sol panel yanında taşma olmaz). Ayrıca görünür yüksekliğe
           göre sınırlanır ki sayfa + kontroller tek ekrana sığsın
           (0.707 = A4 en/boy; 280px ≈ sayfa başlığı + kart dolgusu + kontrol
           satırı + ipucu). Çok kısa ekranlarda 300px tabanın altına inmez. */
        .flip-page {
          flex: 0 1 clamp(300px, calc((100vh - var(--topbar-h) - 280px) * 0.707), 620px);
          min-width: 0;
          max-width: 100%;
          /* PDF çizilene kadar görünen zemin; pdf.js kâğıdı kendisi beyaza boyar. */
          background: var(--surface-3);
          box-shadow: var(--shadow-md);
          border-radius: 2px;
        }
        /* Yüklenirken iskelet A4 oranında yer tutar (canvas boyutlanınca zıplama olmaz). */
        .flip-page.skeleton { aspect-ratio: 210 / 297; }
        .flip-page.empty { visibility: hidden; }
        .flip-page canvas { display: block; width: 100%; height: auto; }
        .flip-anim-next .flip-page.right { animation: flipInR 0.55s ease; transform-origin: left center; }
        .flip-anim-prev .flip-page.left { animation: flipInL 0.55s ease; transform-origin: right center; }
        @keyframes flipInR {
          from { transform: rotateY(-70deg); opacity: 0.4; }
          to { transform: rotateY(0deg); opacity: 1; }
        }
        @keyframes flipInL {
          from { transform: rotateY(70deg); opacity: 0.4; }
          to { transform: rotateY(0deg); opacity: 1; }
        }
        /* Telefon: yan yana iki sayfa okunmaz; sayfalar alt alta, tam genişlik.
           Boş taraf yer kaplamasın diye tamamen gizlenir. */
        @media (max-width: 640px) {
          .flip-stage { flex-direction: column; align-items: center; gap: 12px; }
          .flip-page { flex: 0 0 auto; width: 100%; }
          .flip-page.empty { display: none; }
        }
        /* Telefon / dokunmatik cihazda klavye ipucunun anlamı yok. */
        @media (max-width: 640px), (hover: none) {
          .flip-hint { display: none; }
        }
      `}</style>

      {!doc ? (
        <p style={{ color: "var(--text-2)", padding: 40 }}>Katalog yükleniyor…</p>
      ) : (
        <>
          <div
            key={animKey}
            className={`flip-stage ${direction === "next" ? "flip-anim-next" : "flip-anim-prev"}`}
            aria-busy={!painted}
          >
            <div
              className={`flip-page left${l ? "" : " empty"}${painted ? "" : " skeleton"}`}
              aria-hidden={!l}
            >
              <canvas ref={leftRef} />
            </div>
            <div
              className={`flip-page right${r ? "" : " empty"}${painted ? "" : " skeleton"}`}
              aria-hidden={!r}
            >
              <canvas ref={rightRef} />
            </div>
          </div>

          <div
            className="row no-print"
            style={{ justifyContent: "center", gap: 12, marginTop: 14 }}
          >
            <button
              type="button"
              className="btn small secondary"
              onClick={() => go(-1)}
              disabled={spread === 0}
            >
              <Icon name="chevron-left" size={16} /> Önceki
            </button>
            <span className="num" style={{ color: "var(--text-2)", fontSize: 13 }}>
              {l && r
                ? `Sayfa ${l}–${r} / ${doc.numPages}`
                : `Sayfa ${l || r || 1} / ${doc.numPages}`}
            </span>
            <button
              type="button"
              className="btn small secondary"
              onClick={() => go(1)}
              disabled={spread >= maxSpread}
            >
              Sonraki <Icon name="chevron-right" size={16} />
            </button>
            <a className="btn small" href={pdfUrl} download>
              <Icon name="download" size={16} /> PDF İndir
            </a>
          </div>
          <p
            className="no-print flip-hint"
            style={{ color: "var(--muted)", fontSize: 12, marginTop: 8 }}
          >
            İpucu: klavye ok tuşlarıyla da sayfa çevirebilirsiniz.
          </p>
        </>
      )}
    </div>
  );
}
