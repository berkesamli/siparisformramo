import ts from "typescript";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const EXT = [".ts", ".tsx", ".mts", ".js", ".mjs", ".json"];

function dosyaBul(p) {
  if (fs.existsSync(p) && fs.statSync(p).isFile()) return p;
  for (const e of EXT) if (fs.existsSync(p + e) && fs.statSync(p + e).isFile()) return p + e;
  for (const e of EXT) { const i = path.join(p, "index" + e); if (fs.existsSync(i)) return i; }
  return null;
}

export async function resolve(spec, ctx, next) {
  if (spec.startsWith("@/")) {
    const f = dosyaBul(path.join(ROOT, spec.slice(2)));
    if (f) return { url: pathToFileURL(f).href, shortCircuit: true };
  }
  if ((spec.startsWith("./") || spec.startsWith("../")) && ctx.parentURL && ctx.parentURL.startsWith("file:")) {
    const f = dosyaBul(path.resolve(path.dirname(fileURLToPath(ctx.parentURL)), spec));
    if (f) return { url: pathToFileURL(f).href, shortCircuit: true };
  }
  return next(spec, ctx);
}

export async function load(url, ctx, next) {
  if (url.startsWith("file:") && /\.(ts|tsx|mts)$/.test(url)) {
    const dosya = fileURLToPath(url);
    const src = fs.readFileSync(dosya, "utf8");
    const out = ts.transpileModule(src, {
      fileName: dosya,
      compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022, jsx: ts.JsxEmit.ReactJSX, esModuleInterop: true, verbatimModuleSyntax: false },
    });
    return { format: "module", source: out.outputText, shortCircuit: true };
  }
  if (url.startsWith("file:") && url.endsWith(".json")) {
    return { format: "json", source: fs.readFileSync(fileURLToPath(url), "utf8"), shortCircuit: true };
  }
  return next(url, ctx);
}
