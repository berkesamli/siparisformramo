import test from "node:test";
import assert from "node:assert/strict";
import { gerekliMetre, subeOner } from "@/lib/uretim/sube-oneri";

test("gerekliMetre: dış ölçü çevresi + fire × adet", () => {
  const m = gerekliMetre({ sku: "x", adet: 2, ozet: "", disMm: { w: 540, h: 740 } });
  // çevre 2×(0.54+0.74)=2.56 + 0.3 fire = 2.86 × 2 = 5.72
  assert.equal(m, 5.72);
  assert.equal(gerekliMetre({ sku: "x", adet: 1, ozet: "" }), 0);
});

test("subeOner: şehir ipucu ve stok", async () => {
  const r = await subeOner([{ sku: "ZZZ-YOK", adet: 1, ozet: "", disMm: { w: 500, h: 500 } }], "İstanbul", "Avcılar");
  assert.equal(r.sube, "istanbul");
  assert.match(r.neden, /bulunamadı/);
  const r2 = await subeOner([], "Ankara", "");
  assert.equal(r2.sube, "ankara");
});
