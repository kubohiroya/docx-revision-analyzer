// スモークテスト: アプリをウィンドウを出さずに起動し、fixtures の docx を開いて
// 各タブの画面を PNG に保存する (DRA_SMOKE_OUT、既定は ./smoke-out)。
//   npm run smoke [-- path/to/file.docx]
import { spawnSync } from "node:child_process";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const dir = path.dirname(fileURLToPath(import.meta.url));
const electron = createRequire(import.meta.url)("electron");
const file = path.resolve(process.argv[2] ?? path.join(dir, "../fixtures/flow-demo.docx"));
const out = path.resolve(process.env.DRA_SMOKE_OUT ?? path.join(dir, "smoke-out"));
const r = spawnSync(electron, [dir], {
  stdio: "inherit",
  env: { ...process.env, DRA_SMOKE_FILE: file, DRA_SMOKE_OUT: out },
  timeout: 120000,
});
process.exit(r.status ?? 1);
