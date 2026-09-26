// Testler için TypeScript yükleyici: node --import ./scripts/register-ts.mjs --test tests/*.test.ts
// "@/…" yolunu depo köküne çözer, .ts/.tsx dosyalarını typescript ile anında derler,
// JSON'u modül olarak verir. Yalnızca test ortamı içindir (Next derlemesinde kullanılmaz).
import { register } from "node:module";
register("./ts-loader.mjs", import.meta.url);
