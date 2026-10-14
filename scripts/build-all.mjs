// `npm run build`: 実行している環境で作れるものを一括で作る (pnpm run build でも同じ)。
//
// 成果物は種類ごとに次の場所にまとまる (dist/ は npm パッケージ、配布物はすべて release/ の下):
//
//   dist/                       lib      ライブラリと CLI (TypeScript → JavaScript。npm パッケージの中身)
//   release/cli/<os>-<arch>/    cli      単体実行ファイル docx-revision-chart など (Bun)
//   release/droplet/mac-<arch>/ droplet  macOS のドロップレット docx-revision-chart.app / -flow.app (macOS のみ)
//   release/desktop/<os>/       desktop  デスクトップアプリ (Electron。インストーラと展開済みのアプリ)
//
// <os> は mac / win / linux、<arch> は x64 / arm64 (electron-builder の呼び方に合わせている)。
// 必要なツールが無い工程 (Bun が無い、macOS 以外でのドロップレットなど) は、理由を表示して飛ばす。
//
// オプション (npm run build -- --skip=desktop のように指定する):
//   --only=lib,cli           指定した工程だけを行う
//   --skip=desktop,droplet   指定した工程を飛ばす
//   --cli=chart,flow         単体実行ファイルを作る CLI を絞る (既定はすべて)
//   --all-platforms          単体実行ファイルを主要な OS / CPU 向けにまとめてクロスビルドする
//   --installers             デスクトップアプリのインストーラ (dmg / zip / NSIS) も作る (時間がかかる)
import { spawnSync } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const STEPS = ["lib", "cli", "droplet", "desktop"];
// [CLI 名, 実行ファイル名]
const CLIS = [
  ["chart", "docx-revision-chart"],
  ["flow", "docx-revision-flow"],
  ["score", "docx-ai-suspicion-score"],
  ["snapshot", "docx-revision-snapshot"],
  ["versions", "docx-revision-versions"],
];
const DROPLET_CLIS = ["chart", "flow"];
// [フォルダ名 (<os>-<arch>), Bun のターゲット]
const PLATFORMS = [
  ["linux-x64", "bun-linux-x64"],
  ["linux-arm64", "bun-linux-arm64"],
  ["mac-x64", "bun-darwin-x64"],
  ["mac-arm64", "bun-darwin-arm64"],
  ["win-x64", "bun-windows-x64"],
];

const args = process.argv.slice(2);
const listArg = (name) =>
  args
    .filter((a) => a.startsWith(`--${name}=`))
    .flatMap((a) => a.slice(name.length + 3).split(","))
    .map((s) => s.trim())
    .filter(Boolean);
const only = listArg("only");
const skip = new Set(listArg("skip"));
const cliFilter = listArg("cli");
const allPlatforms = args.includes("--all-platforms");
const installers = args.includes("--installers");
for (const s of [...only, ...skip]) {
  if (!STEPS.includes(s)) {
    console.error(`unknown step "${s}" (steps: ${STEPS.join(", ")})`);
    process.exit(2);
  }
}
for (const c of cliFilter) {
  if (!CLIS.some(([cli]) => cli === c)) {
    console.error(`unknown CLI "${c}" (CLIs: ${CLIS.map(([cli]) => cli).join(", ")})`);
    process.exit(2);
  }
}
const wanted = (step) => (only.length ? only.includes(step) : !skip.has(step));

const isWin = process.platform === "win32";
const hostOs = { darwin: "mac", win32: "win", linux: "linux" }[process.platform] ?? process.platform;
const hostPlatform = `${hostOs}-${process.arch}`;
const exe = (os, name) => (os === "win" ? `${name}.exe` : name);
const rel = (p) => path.relative(root, p).split(path.sep).join("/");
const release = path.join(root, "release");
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

/** Bun で CLI を単体実行ファイルにする。target が無ければこの OS / CPU 向け */
function compileCli(cli, name, platform, target) {
  const dir = path.join(release, "cli", platform);
  fs.mkdirSync(dir, { recursive: true });
  const out = path.join(dir, exe(platform.split("-")[0], name));
  // Bun はコンパイル中の一時ファイル (.<hash>.bun-build) を作業フォルダに作るため、出力先で実行して後で消す
  run("bun", ["build", path.join(root, "src", "cli", `${cli}.ts`), "--compile", ...(target ? [`--target=${target}`] : []), "--outfile", out], {
    cwd: dir,
  });
  for (const f of fs.readdirSync(dir)) if (f.endsWith(".bun-build")) fs.rmSync(path.join(dir, f), { force: true });
  return out;
}

// 1. ライブラリと CLI
step("lib", "Library and CLIs (TypeScript → dist/)", () => {
  fs.rmSync(path.join(root, "dist"), { recursive: true, force: true });
  run(process.execPath, [path.join(root, "node_modules", "typescript", "bin", "tsc"), "-p", "tsconfig.json"]);
  return ["dist/index.js", "dist/core.js", "dist/cli/{chart,flow,score,snapshot,versions}.js"];
});

// 2. 単体実行ファイル
const builtHostCli = new Set();
step("cli", `Single executables (Bun → release/cli/${allPlatforms ? "<os>-<arch>" : hostPlatform}/)`, () => {
  if (!has("bun")) return "Bun isn't installed (https://bun.sh: curl -fsSL https://bun.sh/install | bash)";
  const targets = allPlatforms ? PLATFORMS : [[hostPlatform, undefined]];
  const made = [];
  for (const [platform, target] of targets) {
    for (const [cli, name] of CLIS) {
      if (cliFilter.length && !cliFilter.includes(cli)) continue;
      made.push(rel(compileCli(cli, name, platform, target)));
      if (platform === hostPlatform) builtHostCli.add(cli);
    }
  }
  return made;
});

// 3. macOS のドロップレット (この Mac 向けの実行ファイルを .app に包む。2. で作っていなければここで作る)
step("droplet", `macOS droplets (release/droplet/${hostPlatform}/*.app)`, () => {
  if (process.platform !== "darwin") return "macOS only";
  if (!has("bun")) return "Bun isn't installed";
  if (!has("osacompile")) return "osacompile isn't available";
  const outDir = path.join(release, "droplet", hostPlatform);
  const made = [];
  for (const cli of DROPLET_CLIS) {
    const name = CLIS.find(([c]) => c === cli)[1];
    const bin = builtHostCli.has(cli) ? path.join(release, "cli", hostPlatform, name) : compileCli(cli, name, hostPlatform);
    run("bash", [path.join(root, "scripts", "make-mac-droplet.sh"), cli, bin, outDir]);
    made.push(rel(path.join(outDir, `${name}.app`)));
  }
  return made;
});

// 4. デスクトップアプリ (出力先は desktop/electron-builder.config.cjs の directories.output = release/desktop/<os>/)
const desktop = path.join(root, "desktop");
step("desktop", `Desktop app (Electron → release/desktop/${hostOs}/)${installers ? ", with installers" : ""}`, () => {
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
  const outDir = path.join(release, "desktop", hostOs);
  // 前回の成果物が混ざらないように、この OS の分だけ消してから作る
  fs.rmSync(outDir, { recursive: true, force: true });
  const builderArgs = installers
    ? [builder, "--config", "electron-builder.config.cjs", platformFlag, "--publish", "never"]
    : [builder, "--config", "electron-builder.config.cjs", platformFlag, "dir", arch, "--publish", "never"];
  run(process.execPath, builderArgs, { cwd: desktop, env });
  // 作ったもの: インストーラ (dmg / exe / zip / AppImage) と、展開済みのアプリ (<os>-<arch>/ や win-unpacked/ の中)
  const made = [];
  for (const f of fs.existsSync(outDir) ? fs.readdirSync(outDir) : []) {
    if (f.startsWith(".") || /\.(ya?ml|blockmap)$/.test(f)) continue;
    const p = path.join(outDir, f);
    if (fs.statSync(p).isDirectory()) {
      for (const g of fs.readdirSync(p)) if (/\.(app|exe)$/i.test(g) || g === "docx-revision-analyzer-desktop") made.push(rel(path.join(p, g)));
    } else made.push(rel(p));
  }
  return made;
});

// まとめ
console.log("\n==================== build summary ====================");
const mark = { ok: "✔", skipped: "–", failed: "✘" };
for (const r of results) {
  console.log(`${mark[r.status]} ${r.title}: ${r.status}`);
  for (const line of r.detail.split("\n")) if (line) console.log(`    ${line}`);
}
if (results.some((r) => r.status === "failed")) process.exit(1);
