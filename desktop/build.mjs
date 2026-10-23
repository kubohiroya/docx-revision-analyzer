// デスクトップアプリをバンドルする (esbuild)。
//  - main / preload: Node.js (Electron) 向け。electron は外部モジュール
//  - renderer: ブラウザ向け。ライブラリのコア (../src/core.ts) を含む
import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const dir = path.dirname(fileURLToPath(import.meta.url));
const out = path.join(dir, "dist");
fs.rmSync(out, { recursive: true, force: true });

const common = { bundle: true, logLevel: "warning", sourcemap: false, legalComments: "none" };
await esbuild.build({
  ...common,
  entryPoints: [path.join(dir, "src/main.ts")],
  outfile: path.join(out, "main.js"),
  platform: "node",
  format: "cjs",
  target: "node22",
  // electron-updater は実行時に node_modules から読む (electron-builder が配布物に含める)
  external: ["electron", "electron-updater", "@azure/msal-node"],
});
await esbuild.build({
  ...common,
  entryPoints: [path.join(dir, "src/preload.ts")],
  outfile: path.join(out, "preload.js"),
  platform: "node",
  format: "cjs",
  target: "node22",
  external: ["electron"],
});
await esbuild.build({
  ...common,
  entryPoints: [path.join(dir, "src/ext-preload.ts")],
  outfile: path.join(out, "ext-preload.js"),
  platform: "node",
  format: "cjs",
  target: "node22",
  external: ["electron"],
});
// 拡張の実行環境: 拡張のエントリを import() で読み込むため ES モジュールにする
await esbuild.build({
  ...common,
  entryPoints: [path.join(dir, "src/ext-host/runtime.ts")],
  outfile: path.join(out, "ext-host/runtime.js"),
  platform: "browser",
  format: "esm",
  target: "chrome130",
});
fs.copyFileSync(path.join(dir, "src/ext-host/index.html"), path.join(out, "ext-host/index.html"));
await esbuild.build({
  ...common,
  entryPoints: [path.join(dir, "src/renderer/app.ts")],
  outfile: path.join(out, "renderer/app.js"),
  platform: "browser",
  format: "iife",
  target: "chrome130",
});
for (const f of ["index.html", "styles.css"]) {
  fs.copyFileSync(path.join(dir, "src/renderer", f), path.join(out, "renderer", f));
}
// 保存時のページを描く pdf.js (pages/pdfHost.ts の非表示のウィンドウ)。必要なファイルだけを写す
const pdfHost = path.join(out, "pdf-host");
fs.mkdirSync(path.join(pdfHost, "cmaps"), { recursive: true });
for (const f of ["index.html", "host.mjs"]) fs.copyFileSync(path.join(dir, "src/pdf-host", f), path.join(pdfHost, f));
const pdfjs = path.join(dir, "node_modules/pdfjs-dist");
for (const f of ["pdf.min.mjs", "pdf.worker.min.mjs"]) fs.copyFileSync(path.join(pdfjs, "build", f), path.join(pdfHost, f));
// 日本語の CID から文字を引く CMap (ToUnicode の無い PDF のため)
for (const f of ["Adobe-Japan1-UCS2.bcmap"]) fs.copyFileSync(path.join(pdfjs, "cmaps", f), path.join(pdfHost, "cmaps", f));
console.log(`built ${path.relative(process.cwd(), out)}`);
