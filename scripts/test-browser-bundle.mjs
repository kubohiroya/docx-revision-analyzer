// コア (src/core.ts) がブラウザ向けバンドルで動くことを確かめるテスト。
//
//  1. esbuild で platform: "browser" としてバンドルする。Node.js の組み込みモジュール (fs, path,
//     child_process など) を import していればここで失敗する。
//  2. バンドルを、Node.js の process / require / Buffer の無い隔離環境 (vm) で実行し、fixtures の docx を
//     Uint8Array で渡して、変更履歴の抽出・位置・分類・chart / flow の SVG・設定の書き換えを行う。
//  3. 同じ処理を Node.js 向けバンドルでも行い、結果 (SVG など) が一致することを確かめる。
//
//   node scripts/test-browser-bundle.mjs
import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as path from "node:path";
import * as vm from "node:vm";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const entry = path.join(root, "src/core.ts");

async function bundle(platform) {
  const r = await esbuild.build({
    entryPoints: [entry],
    bundle: true,
    platform,
    format: "iife",
    globalName: "DRA",
    write: false,
    metafile: true,
    logLevel: "silent",
  });
  return { code: r.outputFiles[0].text, inputs: Object.keys(r.metafile.inputs) };
}

/** 隔離環境で実行するテスト本体 (文字列にして渡す)。結果を JSON にできる形で返す */
const scenario = `(async () => {
  DRA.setLang("en");
  const out = {};
  for (const [name, bytes] of Object.entries(FIXTURES)) {
    const input = new Uint8Array(bytes);
    const data = await DRA.extractRevisions(input);
    const model = await DRA.parseDocxLayout(input.buffer);
    const positioned = DRA.extractRevisionPositions(model);
    const rules = DRA.defaultRuleSet(DRA.DEFAULT_BULK_CHARS);
    const ctx = DRA.createAnalysisContext(model, positioned, 1);
    const cls = await DRA.runClassifiers(DRA.defaultClassifiers(rules), ctx);
    DRA.applyClassification(data.events, cls);
    const chart = data.events.length
      ? DRA.renderRevisionChart(DRA.buildBuckets(data.events, data.baselineCharCount, "auto"), {
          title: name,
          categories: cls.categories,
        })
      : "";
    const flow = DRA.buildFlow(model, {
      gapThresholdHours: 1,
      bulkChars: DRA.DEFAULT_BULK_CHARS,
      highlightOf: DRA.highlightCategoryMap(cls),
    });
    const flowSvg = flow.sessions.length ? DRA.renderFlowSvg(flow, { title: name, categories: cls.categories }) : "";
    const preserved = await DRA.preserveHistoryInDocx(input);
    out[name] = {
      events: data.events.length,
      positioned: positioned.length,
      highlights: cls.highlights.length,
      settings: data.settings,
      preserveChanged: preserved.changed,
      chart,
      flowSvg,
    };
  }
  return JSON.stringify(out);
})()`;

const fixtures = {};
for (const f of fs.readdirSync(path.join(root, "fixtures")).filter((f) => f.endsWith(".docx"))) {
  fixtures[f] = [...fs.readFileSync(path.join(root, "fixtures", f))];
}

function fail(msg) {
  console.error(`FAIL: ${msg}`);
  process.exit(1);
}

// 1. ブラウザ向けバンドル
let browser;
try {
  browser = await bundle("browser");
} catch (err) {
  fail(`the core can't be bundled for the browser:\n${err.errors?.map((e) => e.text).join("\n") ?? err}`);
}
const nodeBuiltins = browser.inputs.filter((i) => /^(node:|fs$|path$|child_process$|os$)/.test(i));
if (nodeBuiltins.length > 0) fail(`the browser bundle includes Node.js modules: ${nodeBuiltins.join(", ")}`);

// 2. Node.js の機能の無い隔離環境で実行する
const sandbox = {
  FIXTURES: fixtures,
  console,
  setTimeout,
  clearTimeout,
  queueMicrotask,
  TextEncoder,
  TextDecoder,
  navigator: { language: "en-US" },
};
const context = vm.createContext(sandbox);
// ブラウザと同じく self / window をグローバルオブジェクトにする (JSZip の setImmediate のポリフィルが使う)
vm.runInContext("globalThis.self = globalThis; globalThis.window = globalThis;", context);
for (const name of ["process", "require", "Buffer", "global", "module"]) {
  if (vm.runInContext(`typeof ${name}`, context) !== "undefined") fail(`sandbox unexpectedly has ${name}`);
}
vm.runInContext(browser.code, context, { filename: "core.browser.js" });
const browserResult = JSON.parse(await vm.runInContext(scenario, context));

// 3. Node.js 向けバンドルと比べる
const nodeBundle = await bundle("node");
const nodeContext = { FIXTURES: fixtures, require: createRequire(import.meta.url), console, process, setTimeout, clearTimeout, setImmediate, TextEncoder, TextDecoder };
vm.createContext(nodeContext);
vm.runInContext(nodeBundle.code, nodeContext, { filename: "core.node.js" });
const nodeResult = JSON.parse(await vm.runInContext(scenario, nodeContext));

let checked = 0;
for (const [name, r] of Object.entries(browserResult)) {
  if (r.events > 0 && !r.chart.startsWith("<svg")) fail(`${name}: no chart SVG`);
  if (r.events > 0 && !r.flowSvg.startsWith("<svg")) fail(`${name}: no flow SVG`);
  if (JSON.stringify(r) !== JSON.stringify(nodeResult[name])) fail(`${name}: browser and Node.js results differ`);
  console.log(
    `ok ${name}: ${r.events} events, ${r.positioned} positioned, ${r.highlights} highlights, ` +
      `chart ${r.chart.length} bytes, flow ${r.flowSvg.length} bytes`
  );
  checked++;
}
if (checked === 0) fail("no fixtures were checked");
console.log(`browser bundle: ${(browser.code.length / 1024).toFixed(0)} KiB, ${checked} fixtures OK`);
