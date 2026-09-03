"use client";

// Canlı çerçeve önizleme — olgacerceve.com sitesindeki hesaplayıcının
// (websitesi-cercevehesaplayici deposu) önizleme sisteminin React portu.
// Kurallar birebir korunur: oda mockup ölçekleme (wallWidthMM + sıkıştırılmış
// üs ölçeği), ışık yönüne göre gölge, 45° bevel, kadife/metalik dokular,
// gerçek çerçeve border-image + clipRatio/outset tekniği, cam katmanları.

import { useEffect, useMemo, useRef, useState } from "react";
import {
  type FrameImage,
  FRAME_SLICE,
  FRAME_BORDER_SCALE,
  DEFAULT_CLIP_RATIO,
} from "@/data/frame-images";

// ---- Sabitler (siteden birebir) ----

const ART_BG_TEXTURE = `
  repeating-linear-gradient(0deg, rgba(160,120,70,0.25) 0px, rgba(200,170,130,0.12) 1px, rgba(140,100,55,0.18) 2px, transparent 3px, transparent 5px),
  repeating-linear-gradient(90deg, rgba(0,0,0,0.03) 0px, rgba(255,255,255,0.02) 1px, transparent 2px),
  #b8976a`;

interface RoomBg {
  id: string;
  label: string;
  icon: string;
  url: string;
  lightFrom: "left" | "right" | "top";
  bgPos: string;
  wallWidthMM: number;
  frameTop: string;
  frameOffsetX: number; // kutu genişliğine göre % kayma
  frameOffsetY: string;
}

const ROOM_BACKGROUNDS: RoomBg[] = [
  {
    id: "living",
    label: "Oturma Odası",
    icon: "🛋",
    url: "https://cdn.myikas.com/images/04a76b35-2c55-499a-b485-0058f5ce13ce/b404a66b-0708-463c-9b4a-c6169134275f/image_1080.webp",
    lightFrom: "left",
    bgPos: "center center",
    wallWidthMM: 2060,
    frameTop: "30%",
    frameOffsetX: -5,
    frameOffsetY: "0",
  },
  {
    id: "bedroom",
    label: "Yatak Odası",
    icon: "🛏",
    url: "https://cdn.myikas.com/images/04a76b35-2c55-499a-b485-0058f5ce13ce/21cf8e6d-f2b0-4037-959e-ca681d511649/image_1080.webp",
    lightFrom: "right",
    bgPos: "center center",
    wallWidthMM: 2100,
    frameTop: "7%",
    frameOffsetX: -22,
    frameOffsetY: "0",
  },
];

function getRoomShadow(lightFrom: string): string {
  switch (lightFrom) {
    case "left":
      return "6px 8px 18px rgba(0,0,0,0.28), 3px 4px 8px rgba(0,0,0,0.15)";
    case "right":
      return "-6px 8px 18px rgba(0,0,0,0.28), -3px 4px 8px rgba(0,0,0,0.15)";
    case "top":
      return "0 10px 22px rgba(0,0,0,0.30), 0 5px 10px rgba(0,0,0,0.15)";
    default:
      return "0 8px 20px rgba(0,0,0,0.30), 0 3px 8px rgba(0,0,0,0.15)";
  }
}

// Paspartu türüne göre doku (dış katman) — tür KODUYLA seçilir (fiyat bayiye göre değişir)
function matBackground(type: string, code: string, hex: string): string {
  const c = (code || "").toUpperCase();
  if (type === "AG") {
    if (c === "W232") {
      return `repeating-linear-gradient(180deg, rgba(255,250,200,0.4) 0px, rgba(160,120,0,0.12) 1px, rgba(255,245,180,0.3) 2px, rgba(255,255,230,0.15) 3px, transparent 4px),
        repeating-linear-gradient(90deg, rgba(255,255,255,0.15) 0px, transparent 1px, rgba(255,255,200,0.08) 2px, transparent 3px),
        #d4af37`;
    }
    if (c === "W233") {
      return `repeating-linear-gradient(180deg, rgba(255,255,255,0.45) 0px, rgba(140,140,140,0.15) 1px, rgba(250,250,250,0.35) 2px, rgba(255,255,255,0.2) 3px, transparent 4px),
        repeating-linear-gradient(90deg, rgba(255,255,255,0.18) 0px, transparent 1px, rgba(230,230,230,0.1) 2px, transparent 3px),
        #c8c8c8`;
    }
    return hex || "#ffffff";
  }
  // Kadife/Premium dokular — kumaş mikro-doku
  if (type === "KDF" || type === "PR") {
    return `repeating-linear-gradient(90deg, rgba(0,0,0,.03) 0px, rgba(255,255,255,.02) 1px, rgba(0,0,0,.02) 2px, transparent 3px),
      repeating-linear-gradient(0deg, rgba(0,0,0,.02) 0px, rgba(255,255,255,.015) 1px, transparent 2px),
      ${hex || "#ffffff"}`;
  }
  return hex || "#ffffff";
}

// İç paspartu (2. katman) dokusu
function mat2Background(type: string, code: string, hex: string): string {
  const c = (code || "").toUpperCase();
  if (type === "AG") {
    if (c === "W232")
      return "linear-gradient(135deg,#fff8d9 0%,#f4d98a 18%,#d4af37 42%,#fff2bf 55%,#b8860b 78%,#fff6cf 100%)";
    if (c === "W233")
      return "linear-gradient(135deg,#ffffff 0%,#e6e6e6 18%,#bfbfbf 42%,#f8f8f8 55%,#9b9b9b 78%,#ffffff 100%)";
    return hex || "#ffffff";
  }
  if (type === "KDF" || type === "PR") {
    return `repeating-linear-gradient(90deg, rgba(0,0,0,.03) 0px, rgba(255,255,255,.02) 1px, rgba(0,0,0,.02) 2px, transparent 3px),
      repeating-linear-gradient(0deg, rgba(0,0,0,.02) 0px, rgba(255,255,255,.015) 1px, transparent 2px),
      radial-gradient(ellipse at 25% 25%, rgba(255,255,255,.08), transparent 60%),
      radial-gradient(ellipse at 75% 75%, rgba(0,0,0,.1), transparent 60%),
      ${hex || "#ffffff"}`;
  }
  return hex || "#ffffff";
}

type GlassId = "none" | "plain" | "mat" | "museum";

function glassIdFromName(name: string): GlassId {
  if (name === "Mat Cam") return "mat";
  if (name === "Müze Camı") return "museum";
  if (name === "Cam Yok" || !name) return "none";
  return "plain"; // Düz Cam, PVC Cam
}

export interface FramePreviewProps {
  wMM: number;
  hMM: number;
  matTop: number;
  matRight: number;
  matBottom: number;
  matLeft: number;
  hasMat: boolean;
  matCode: string;
  matName: string;
  matColorCode: string;
  matColorHex: string;
  doubleMat: boolean;
  innerMatCode: string;
  innerMatName: string;
  innerColorCode: string;
  innerColorHex: string;
  mountingMM: number;
  zeminEnabled: boolean;
  zeminColorHex: string;
  glassName: string;
  frameImg: (FrameImage & { sku?: string }) | null;
  fullCode: string;
  artImageUrl: string | null;
}

export default function FramePreview(p: FramePreviewProps) {
  const boxRef = useRef<HTMLDivElement | null>(null);
  const [boxSize, setBoxSize] = useState({ w: 0, h: 0 });
  const [roomId, setRoomId] = useState<string>("plain");
  const [artZoom, setArtZoom] = useState(1);
  const [artPosX, setArtPosX] = useState(50);
  const [artPosY, setArtPosY] = useState(50);
  const [artNatural, setArtNatural] = useState<{ w: number; h: number } | null>(null);
  const [frameImgOk, setFrameImgOk] = useState(false);

  // Çerçeve görselini önden yükle — yüklenemezse koyu düz çerçeveye düşülür
  useEffect(() => {
    setFrameImgOk(false);
    if (!p.frameImg?.url) return;
    let cancelled = false;
    const img = new Image();
    img.onload = () => {
      if (!cancelled) setFrameImgOk(true);
    };
    img.src = p.frameImg.url;
    return () => {
      cancelled = true;
    };
  }, [p.frameImg?.url]);

  useEffect(() => {
    const el = boxRef.current;
    if (!el) return;
    const ro = new ResizeObserver((entries) => {
      const r = entries[0]?.contentRect;
      if (r) setBoxSize({ w: r.width, h: r.height });
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // Eser görselinin doğal oranı (zoom hesabı için)
  useEffect(() => {
    setArtZoom(1);
    setArtPosX(50);
    setArtPosY(50);
    setArtNatural(null);
    if (!p.artImageUrl) return;
    const img = new Image();
    img.onload = () => setArtNatural({ w: img.naturalWidth, h: img.naturalHeight });
    img.src = p.artImageUrl;
  }, [p.artImageUrl]);

  function artBg(cW: number, cH: number): string | null {
    if (!p.artImageUrl) return null;
    if (!artNatural || !cW || !cH) {
      return `url('${p.artImageUrl}') ${artPosX}% ${artPosY}% / cover no-repeat`;
    }
    const imgAR = artNatural.w / artNatural.h;
    const cAR = cW / cH;
    const coverPct = Math.max(100, (100 * imgAR) / cAR);
    const zPct = Math.round(coverPct * artZoom);
    return `url('${p.artImageUrl}') ${artPosX}% ${artPosY}% / ${zPct}% auto no-repeat`;
  }

  const room = ROOM_BACKGROUNDS.find((r) => r.id === roomId) || null;
  const isRoomMode = Boolean(room);
  const hasDims = p.wMM > 0 && p.hMM > 0;
  const hasRealFrame = Boolean(p.frameImg) && frameImgOk;
  const bare = Boolean(p.frameImg?.bareFrame) && frameImgOk;
  const glassId = bare ? "none" : glassIdFromName(p.glassName);
  const isDouble = p.doubleMat && p.hasMat && !bare;
  const bevelPx = 2;

  const calc = useMemo(() => {
    const boxW = boxSize.w;
    const boxH = boxSize.h;
    const clipRatio = p.frameImg?.clipRatio ?? DEFAULT_CLIP_RATIO;
    const slice = p.frameImg?.slice ?? FRAME_SLICE;

    // ---- Varsayılan durum (ölçü yok) ----
    if (!hasDims || boxW < 50 || boxH < 50) {
      const defaultBorderPx = hasRealFrame
        ? Math.max(15, Math.round(35 * FRAME_BORDER_SCALE))
        : 15;
      return {
        frameW: 200,
        frameH: 200,
        borderPx: hasRealFrame ? defaultBorderPx : 0,
        paddingPx: hasRealFrame ? 0 : defaultBorderPx,
        outset: 0,
        slice,
        matPad: { t: 15, r: 15, b: 15, l: 15 },
        showMat: !bare,
        artFill: false,
        mountingPx: 0,
        outerPad: { t: 15, r: 15, b: 15, l: 15 },
        isDefault: true,
      };
    }

    // ---- Ölçülü durum ----
    const safePad = 6;
    let availW = Math.max(90, boxW - safePad * 2);
    let availH = Math.max(90, boxH - safePad * 2);

    const totalW = Math.max(p.wMM + p.matLeft + p.matRight, p.wMM);
    const totalH = Math.max(p.hMM + p.matTop + p.matBottom, p.hMM);

    // Oda modu: duvar genişliğine göre, üs ile sıkıştırılmış gerçekçi ölçek
    if (isRoomMode && room) {
      const virtualWallMM = room.wallWidthMM || 2400;
      const refMM = 700;
      const compressionExp = 0.62;
      const pxPerMM = availW / virtualWallMM;
      const maxMM = Math.max(totalW, totalH);
      const compressedFactor = Math.pow(maxMM / refMM, compressionExp - 1);
      let targetW = totalW * pxPerMM * compressedFactor;
      let targetH = totalH * pxPerMM * compressedFactor;
      const capScale = Math.min(1, (availW * 0.5) / targetW, (availH * 0.45) / targetH);
      targetW *= capScale;
      targetH *= capScale;
      if (Math.min(targetW, targetH) < 55) {
        const up = 55 / Math.min(targetW, targetH);
        targetW *= up;
        targetH *= up;
      }
      availW = targetW;
      availH = targetH;
    }

    // Perspektif: küçük eser → kalın kenar, büyük eser → ince kenar
    const maxDimMM = Math.max(totalW, totalH);
    let frameBorderPx: number;
    if (isRoomMode) {
      const roomBorderBase = 12;
      frameBorderPx = hasRealFrame
        ? Math.round(roomBorderBase * FRAME_BORDER_SCALE)
        : roomBorderBase;
    } else {
      const perspectiveFactor = Math.max(0.6, Math.min(1.4, 400 / maxDimMM));
      const base = Math.max(
        16,
        Math.min(55, Math.round(Math.min(availW, availH) * 0.16 * perspectiveFactor))
      );
      frameBorderPx = hasRealFrame
        ? Math.max(8, Math.round(base * FRAME_BORDER_SCALE))
        : base;
    }

    const innerMin = isRoomMode ? 20 : 40;
    const innerW = Math.max(innerMin, availW - frameBorderPx * 2);
    const innerH = Math.max(innerMin, availH - frameBorderPx * 2);
    const scale = Math.min(innerW / totalW, innerH / totalH);
    const contentW = Math.max(30, totalW * scale);
    const contentH = Math.max(30, totalH * scale);

    // Paspartu kenarları (px) — min 8px görünürlük
    const minMatPx = 8;
    const hasMatEdges =
      p.hasMat &&
      !bare &&
      (p.matTop > 0 || p.matBottom > 0 || p.matLeft > 0 || p.matRight > 0);
    const clampEdge = (mm: number, limit: number) =>
      hasMatEdges && mm > 0
        ? Math.max(minMatPx, Math.min(mm * scale, (limit - 20) / 2))
        : 0;
    const cTop = clampEdge(p.matTop, contentH);
    const cBottom = clampEdge(p.matBottom, contentH);
    const cLeft = clampEdge(p.matLeft, contentW);
    const cRight = clampEdge(p.matRight, contentW);

    // artFill: paspartusuz + eser görselli + gerçek çerçeveli →
    // görsel çerçevenin altına uzanır, clipRatio/outset ile az kesilir
    const artFill = Boolean(p.artImageUrl) && !hasMatEdges && hasRealFrame;

    let borderPx = hasRealFrame ? frameBorderPx : 0;
    let paddingPx = hasRealFrame ? 0 : frameBorderPx;
    let outset = 0;
    let frameW = contentW + frameBorderPx * 2;
    let frameH = contentH + frameBorderPx * 2;
    if (artFill) {
      const cssBorder = Math.round(frameBorderPx * clipRatio);
      outset = frameBorderPx - cssBorder;
      borderPx = cssBorder;
      frameW = contentW + cssBorder * 2;
      frameH = contentH + cssBorder * 2;
    }

    // Çift paspartu iç katman
    const mountingPx = Math.max(5, (p.mountingMM || 5) * scale);
    const minOuterPx = 6;
    const outerPad = isDouble
      ? {
          t: Math.max(minOuterPx, cTop - mountingPx - bevelPx),
          r: Math.max(minOuterPx, cRight - mountingPx - bevelPx),
          b: Math.max(minOuterPx, cBottom - mountingPx - bevelPx),
          l: Math.max(minOuterPx, cLeft - mountingPx - bevelPx),
        }
      : {
          t: Math.max(0, cTop - bevelPx),
          r: Math.max(0, cRight - bevelPx),
          b: Math.max(0, cBottom - bevelPx),
          l: Math.max(0, cLeft - bevelPx),
        };

    return {
      frameW,
      frameH,
      borderPx,
      paddingPx,
      outset,
      slice,
      showMat: hasMatEdges,
      artFill,
      mountingPx,
      outerPad,
      contentW,
      contentH,
      cTop,
      cBottom,
      cLeft,
      cRight,
      isDefault: false,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    boxSize, hasDims, hasRealFrame, bare, isRoomMode, room, isDouble,
    p.wMM, p.hMM, p.matTop, p.matRight, p.matBottom, p.matLeft,
    p.hasMat, p.mountingMM, p.frameImg, p.artImageUrl,
  ]);

  // ---- Stiller ----
  const outerMatBg = matBackground(p.matCode, p.matColorCode, p.matColorHex);
  const innerMatBg = mat2Background(p.innerMatCode, p.innerColorCode, p.innerColorHex);

  const artAreaW = (calc.contentW || 170) - (calc.cLeft || 0) - (calc.cRight || 0);
  const artAreaH = (calc.contentH || 170) - (calc.cTop || 0) - (calc.cBottom || 0);
  const artBackground = bare
    ? "transparent"
    : calc.artFill
      ? "transparent"
      : p.artImageUrl
        ? artBg(artAreaW, artAreaH) || `url('${p.artImageUrl}') center/cover no-repeat`
        : p.zeminEnabled && p.zeminColorHex && p.zeminColorHex !== "-"
          ? p.zeminColorHex
          : ART_BG_TEXTURE;

  const frameStyle: React.CSSProperties = {
    width: calc.frameW,
    height: calc.frameH,
    padding: calc.paddingPx,
    borderStyle: "solid",
    borderWidth: calc.borderPx,
    // Görsel yüklenemezse koyu çerçeve rengi görünsün (border-image yüklenince
    // border-color'ın üzerine çizilir)
    borderColor: "#241a10",
    boxShadow: isRoomMode && room ? getRoomShadow(room.lightFrom) : "0 8px 32px rgba(0,0,0,0.3)",
    transform:
      isRoomMode && room
        ? `translate(${((room.frameOffsetX || 0) / 100) * (boxSize.w || 300)}px, ${room.frameOffsetY || 0})`
        : undefined,
  };
  if (hasRealFrame && p.frameImg) {
    frameStyle.borderImageSource = `url('${p.frameImg.url}')`;
    frameStyle.borderImageSlice = calc.slice;
    frameStyle.borderImageRepeat = "stretch";
    frameStyle.borderImageOutset = calc.outset ? `${calc.outset}px` : 0;
    if (bare) {
      frameStyle.background = "transparent";
    } else if (calc.showMat) {
      frameStyle.background = outerMatBg;
    } else if (calc.artFill) {
      frameStyle.background =
        artBg(calc.frameW, calc.frameH) || `url('${p.artImageUrl}') center/cover no-repeat`;
      frameStyle.backgroundOrigin = "border-box";
    } else if (p.artImageUrl && calc.isDefault) {
      frameStyle.background =
        artBg(200, 200) || `url('${p.artImageUrl}') center/cover no-repeat`;
      frameStyle.backgroundOrigin = "border-box";
    } else {
      frameStyle.background = ART_BG_TEXTURE;
    }
  } else {
    frameStyle.background = "#141210";
  }

  const matOuterStyle: React.CSSProperties = calc.showMat
    ? {
        padding: `${calc.outerPad.t}px ${calc.outerPad.r}px ${calc.outerPad.b}px ${calc.outerPad.l}px`,
        background: outerMatBg,
      }
    : calc.isDefault && !bare
      ? { padding: 15, background: hasRealFrame ? "transparent" : ART_BG_TEXTURE }
      : { padding: 0, background: "transparent" };

  const bevelOuterStyle: React.CSSProperties = calc.showMat
    ? { padding: bevelPx, background: "#ffffff" }
    : { padding: 0, background: "transparent" };

  const glassInset =
    glassId !== "none" ? (calc.borderPx > 0 ? `-${calc.borderPx}px` : "0") : "0";

  // Çerçeve dokusu overlay: cam parlamasının ÜSTÜNDE kalır (oda modunda gizli)
  const showFrameImage = hasRealFrame && !isRoomMode && Boolean(p.frameImg);

  const label = hasDims
    ? `${p.wMM + p.matLeft + p.matRight}×${p.hMM + p.matTop + p.matBottom} mm`
    : "";

  let noteText: string;
  if (!hasDims) {
    noteText = "Ölçü girince canlı olarak güncellenir.";
  } else {
    let matTxt: string;
    if (p.hasMat) {
      matTxt = isDouble
        ? `Dış: ${p.matName} (${p.matColorCode || "-"}) | İç: ${p.innerMatName} (${p.innerColorCode || "-"})`
        : `Paspartu: ${p.matName}${p.matColorCode ? ` (${p.matColorCode})` : ""}`;
    } else {
      matTxt = "Paspartu: Yok";
    }
    noteText = `${matTxt} • Cam: ${p.glassName || "Cam Yok"}`;
  }

  return (
    <div className="fp">
      <div className="fp-head">
        <b>Önizleme</b>
        <span className="fp-label">{label}</span>
        <span style={{ flex: 1 }} />
        <div className="fp-rooms">
          <button
            type="button"
            className={roomId === "plain" ? "active" : ""}
            onClick={() => setRoomId("plain")}
          >
            Sade
          </button>
          {ROOM_BACKGROUNDS.map((r) => (
            <button
              key={r.id}
              type="button"
              className={roomId === r.id ? "active" : ""}
              onClick={() => setRoomId(r.id)}
              title={r.label}
            >
              {r.icon}
            </button>
          ))}
        </div>
      </div>

      <div
        ref={boxRef}
        className={`fp-box ${isRoomMode ? "room" : ""}`}
        style={
          isRoomMode && room
            ? {
                background: `url('${room.url}') ${room.bgPos} / cover no-repeat`,
                paddingTop: room.frameTop,
              }
            : undefined
        }
      >
        {!hasDims && (
          <div className="fp-empty">
            <div className="fp-empty-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="var(--brand)" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                <line x1="3" y1="9" x2="21" y2="9" />
                <line x1="3" y1="15" x2="21" y2="15" />
                <line x1="9" y1="3" x2="9" y2="21" />
                <line x1="15" y1="3" x2="15" y2="21" />
              </svg>
            </div>
            <div className="fp-empty-text">
              Eser ölçüsünü girin,
              <br />
              <strong>canlı önizleme</strong> burada görünecek.
            </div>
          </div>
        )}

        {hasDims && (
          <div className="fp-frame-wrapper">
            <div className="fp-frame" style={frameStyle}>
              <div className="fp-mat-outer" style={matOuterStyle}>
                <div className="fp-bevel-outer" style={bevelOuterStyle}>
                  {isDouble ? (
                    <div
                      className="fp-mat-inner"
                      style={{
                        padding: Math.max(3, calc.mountingPx - bevelPx),
                        background: innerMatBg,
                      }}
                    >
                      <div className="fp-bevel-inner" style={{ padding: bevelPx, background: "#fff" }}>
                        <div className="fp-art" style={{ background: artBackground }} />
                      </div>
                    </div>
                  ) : (
                    <div className="fp-art" style={{ background: artBackground }} />
                  )}
                </div>
              </div>
              {glassId !== "none" && (
                <div className={`fp-glass fp-glass--${glassId}`} style={{ inset: glassInset }} />
              )}
              {showFrameImage && p.frameImg && (
                <div
                  className="fp-frame-image"
                  style={{
                    width: calc.frameW,
                    height: calc.frameH,
                    borderWidth: calc.borderPx,
                    borderImageSource: `url('${p.frameImg.url}')`,
                    borderImageSlice: calc.slice,
                    borderImageRepeat: "stretch",
                    borderImageOutset: calc.outset ? `${calc.outset}px` : 0,
                  }}
                />
              )}
            </div>
          </div>
        )}
      </div>

      {p.artImageUrl && hasDims && (
        <div className="fp-art-controls">
          <label>
            🔍
            <input
              type="range" min="1" max="2.5" step="0.05"
              value={artZoom}
              onChange={(e) => setArtZoom(parseFloat(e.target.value))}
            />
          </label>
          <label>
            ↔
            <input
              type="range" min="0" max="100" step="1"
              value={artPosX}
              onChange={(e) => setArtPosX(parseInt(e.target.value))}
            />
          </label>
          <label>
            ↕
            <input
              type="range" min="0" max="100" step="1"
              value={artPosY}
              onChange={(e) => setArtPosY(parseInt(e.target.value))}
            />
          </label>
        </div>
      )}

      {hasRealFrame ? (
        <p className="fp-note ok">✓ {p.frameImg?.sku || p.fullCode} — gerçek profil görseli</p>
      ) : p.fullCode ? (
        <p className="fp-note">{p.fullCode} için görsel henüz eklenmedi</p>
      ) : null}
      <p className="fp-note">{noteText}</p>
    </div>
  );
}
