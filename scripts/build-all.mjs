// `npm run build`: 実行している環境で作れるものを一括で作る (pnpm run build でも同じ)。
//
//   1. lib       ライブラリと CLI (TypeScript → dist/)
//   2. binaries  単体実行ファイル docx-revision-chart / -flow / docx-ai-suspicion-score (Bun → dist-bin/)
//   3. droplets  macOS のドロップレット docx-revision-chart.app / -flow.app (dist-bin/。macOS のみ)
//   4. desktop   デスクトップアプリ (Electron → desktop/release/)。この OS / CPU 向け、署名なし
//
// 必要なツールが無い工程 (Bun が無い、macOS 以外でのドロップレットなど) は、理由を表示して飛ばす。
//
// オプション (npm run build -- --skip=desktop のように指定する):
//   --only=lib,binaries      指定した工程だけを行う
//   --skip=desktop,droplets  指定した工程を飛ばす
//   --all-platforms          単体実行ファイルを主要な OS / CPU 向けにまとめてクロスビルドする
//   --installers             デスクトップアプリのインストーラ (dmg / zip / NSIS) も作る (時間がかかる)
//
// 個別のコマンド (build:lib, build:binary:flow, build:mac-app:flow など) もこれまでどおり使える。
import { spawnSync } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const STEPS = ["lib", "binaries", "droplets", "desktop"];

const args = process.argv.slice(2);
const listArg = (name) =>
  args
    .filter((a) => a.startsWith(`--${name}=`))
    .flatMap((a) => a.slice(name.length + 3).split(","))
    .map((s) => s.trim())
    .filter(Boolean);
const only = listArg("only");
const skip = new Set(listArg("skip"));
const allPlatforms = args.includes("--all-platforms");
const installers = args.includes("--installers");
for (const s of [...only, ...skip]) {
  if (!STEPS.includes(s)) {
    console.error(`unknown step "${s}" (steps: ${STEPS.join(", ")})`);
    process.exit(2);
  }
}
const wanted = (step) => (only.length ? only.includes(step) : !skip.has(step));

const isWin = process.platform === "win32";
const results = [];

function has(cmd) {
  const r = spawnSync(isWin ? "where" : "which", [cmd], { stdio: "ignore" });
  return r.status === 0;
}

function run(cmd, cmdArgs, opts = {}) {
  console.log(`\n$ ${[cmd, ...cmdArgs].join(" ")}`);
  const r = spawnSync(cmd, cmdArgs, { stdio: "inherit", cwd: root, shell: isWin, ...opts });
  if (r.status !== 0) throw new Error(`${cmd} ${cmdArgs.join(" ")} failed (exit ${r.status ?? r.signal})`);
}

/** 工程を実行し、結果 (作成したもの / 飛ばした理由 / 失敗) を記録する */
function step(name, title, fn) {
  if (!wanted(name)) {
    results.push({ name, title, status: "skipped", detail: "not selected" });
    return;
  }
  console.log(`\n==> ${title}`);
  try {
    const out = fn();
    if (typeof out === "string") results.push({ name, title, status: "skipped", detail: out });
    else results.push({ name, title, status: "ok", detail: out.join("\n") });
  } catch (err) {
    results.push({ name, title, status: "failed", detail: err instanceof Error ? err.message : String(err) });
  }
}

// 1. ライブラリと CLI
step("lib", "Library and CLIs (TypeScript → dist/)", () => {
  fs.rmSync(path.join(root, "dist"), { recursive: true, force: true });
  run(process.execPath, [path.join(root, "node_modules", "typescript", "bin", "tsc"), "-p", "tsconfig.json"]);
  return ["dist/index.js", "dist/core.js", "dist/cli/{chart,flow,score}.js"];
});

// 2. 単体実行ファイル
const CLIS = [
  ["chart", "docx-revision-chart"],
  ["flow", "docx-revision-flow"],
  ["score", "docx-ai-suspicion-score"],
];
const binDir = path.join(root, "dist-bin");
step("binaries", "Single executables (Bun → dist-bin/)", () => {
  if (!has("bun")) return "Bun isn't installed (https://bun.sh: curl -fsSL https://bun.sh/install | bash)";
  fs.mkdirSync(binDir, { recursive: true });
  const targets = allPlatforms
    ? [
        ["bun-linux-x64", "-linux-x64"],
        ["bun-linux-arm64", "-linux-arm64"],
        ["bun-darwin-x64", "-macos-x64"],
        ["bun-darwin-arm64", "-macos-arm64"],
        ["bun-windows-x64", "-windows-x64.exe"],
      ]
    : [[undefined, isWin ? ".exe" : ""]];
  const made = [];
  for (const [cli, name] of CLIS) {
    for (const [target, suffix] of targets) {
      const out = path.join(binDir, `${name}${suffix}`);
      // Bun はコンパイル中の一時ファイル (.<hash>.bun-build) を作業フォルダに作るため、dist-bin/ で実行する
      run("bun", ["build", path.join(root, "src", "cli", `${cli}.ts`), "--compile", ...(target ? [`--target=${target}`] : []), "--outfile", out], {
        cwd: binDir,
      });
      made.push(path.relative(root, out));
    }
  }
  // Bun が残す一時ファイルを消す
  for (const f of fs.readdirSync(binDir)) if (f.endsWith(".bun-build")) fs.rmSync(path.join(binDir, f), { force: true });
  return made;
});

// 3. macOS のドロップレット (2. で作ったこの Mac 向けの実行ファイルを使う)
step("droplets", "macOS droplets (dist-bin/*.app)", () => {
  if (process.platform !== "darwin") return "macOS only";
  if (!has("bun")) return "Bun isn't installed";
  if (!has("osacompile")) return "osacompile isn't available";
  const made = [];
  for (const cli of ["chart", "flow"]) {
    const bin = path.join(binDir, cli === "chart" ? "docx-revision-chart" : "docx-revision-flow");
    // 実行ファイルが既にあれば作り直さない (SKIP_BINARY_BUILD)
    run("bash", [path.join(root, "scripts", "make-mac-droplet.sh"), cli], {
      env: { ...process.env, ...(fs.existsSync(bin) ? { SKIP_BINARY_BUILD: "1" } : {}) },
    });
    made.push(`dist-bin/${path.basename(bin)}.app`);
  }
  return made;
});

// 4. デスクトップアプリ
const desktop = path.join(root, "desktop");
step("desktop", `Desktop app (Electron → desktop/release/)${installers ? ", with installers" : ""}`, () => {
  if (!fs.existsSync(path.join(desktop, "package.json"))) return "desktop/ not found";
  // デスクトップアプリの依存 (Electron など) が無ければ入れる (desktop/package-lock.json に従う)
  if (!fs.existsSync(path.join(desktop, "node_modules", "electron-builder"))) {
    run("npm", ["install", "--no-audit", "--no-fund"], { cwd: desktop });
  }
  run(process.execPath, [path.join(desktop, "build.mjs")], { cwd: desktop });
  const platformFlag = { darwin: "--mac", win32: "--win", linux: "--linux" }[process.platform];
  if (!platformFlag) return `packaging isn't supported on ${process.platform}`;
  const arch = { arm64: "--arm64", x64: "--x64" }[process.arch] ?? "--x64";
  const builder = path.join(desktop, "node_modules", "electron-builder", "cli.js");
  // 証明書が無ければ署名しない (キーチェーンの証明書を探しに行かない)
  const env = { ...process.env };
  if (!env.CSC_LINK && !env.CSC_NAME) env.CSC_IDENTITY_AUTO_DISCOVERY = "false";
  const builderArgs = installers
    ? [builder, "--config", "electron-builder.config.cjs", platformFlag, "--publish", "never"]
    : [builder, "--config", "electron-builder.config.cjs", platformFlag, "dir", arch, "--publish", "never"];
  run(process.execPath, builderArgs, { cwd: desktop, env });
  // 作ったもの: インストーラ (dmg / exe / zip / AppImage) と、展開済みのアプリ (release/<os>-<arch>/ の中)
  const release = path.join(desktop, "release");
  const made = [];
  for (const f of fs.readdirSync(release)) {
    if (f.startsWith(".") || /\.(ya?ml|blockmap)$/.test(f)) continue;
    const p = path.join(release, f);
    if (fs.statSync(p).isDirectory()) {
      for (const g of fs.readdirSync(p)) if (/\.(app|exe)$/i.test(g) || g === "resources") made.push(`desktop/release/${f}/${g}`);
    } else made.push(`desktop/release/${f}`);
  }
  return made.filter((m) => !m.endsWith("/resources")).length ? made.filter((m) => !m.endsWith("/resources")) : made;
});

// まとめ
console.log("\n==================== build summary ====================");
const mark = { ok: "✔", skipped: "–", failed: "✘" };
for (const r of results) {
  console.log(`${mark[r.status]} ${r.title}: ${r.status}`);
  for (const line of r.detail.split("\n")) if (line) console.log(`    ${line}`);
}
if (results.some((r) => r.status === "failed")) process.exit(1);
