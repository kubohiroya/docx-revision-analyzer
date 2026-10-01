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
  external: ["electron", "electron-updater"],
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
console.log(`built ${path.relative(process.cwd(), out)}`);
