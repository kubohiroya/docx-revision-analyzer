// スナップショットの保管と通し解析 (src/lib/snapshots.ts, src/node/snapshotArchive.ts, src/cli/snapshot.ts) のテスト。
//
// 区切り (スプリント) ごとの提出を JSZip で作り、docx-revision-snapshot を実際に動かして、通し番号での保管・
// 同じ内容の二重保管の防止・各回で新しく加わった変更の数・改ざんの痕跡・出力ファイル名・一覧を確かめる。
//
//   1. 第1回: テンプレート (ロック済み) に、記録しながら書き足したもの
//   2. 第2回: 教員が第1回をすべて承諾してロックをかけ直して返したものに、書き足し・承諾済みの文章の削除をしたもの
//   3. 第3回: 第2回を承諾せずに返したものに書き足したもの。ただし記録をオフにして入力した段落がある
//   4. 第4回: 第3回より前の日時の変更が初めて現れる (別の場所で編集したファイルに差し替えた)
//
//   node scripts/test-snapshots.mjs            (KEEP_TMP=1 で出力を残す)
import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const { XMLValidator } = require("fast-xml-parser");

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-snapshots-"));
const cli = path.join(tmp, "snapshot.cjs");
await esbuild.build({
  entryPoints: [path.join(root, "src/cli/snapshot.ts")],
  bundle: true,
  platform: "node",
  format: "cjs",
  outfile: cli,
  logLevel: "silent",
});

let failed = 0;
function check(name, cond, detail = "") {
  if (cond) console.log(`ok ${name}`);
  else {
    failed++;
    console.error(`FAIL ${name}${detail ? `: ${detail}` : ""}`);
  }
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const LOCK = (salt) =>
  `<w:documentProtection w:edit="trackedChanges" w:enforcement="1" w:cryptAlgorithmSid="14" w:hash="HASH-${salt}" w:salt="${salt}"/>`;
function settings({ rsids, track = true, salt = "S0" }) {
  return (
    `<?xml version="1.0"?><w:settings ${W}>${track ? "<w:trackRevisions/>" : ""}${LOCK(salt)}` +
    `<w:rsids><w:rsidRoot w:val="00A00001"/>${rsids.map((r) => `<w:rsid w:val="${r}"/>`).join("")}</w:rsids></w:settings>`
  );
}
const run = (rsid, text) => `<w:r w:rsidR="${rsid}"><w:t xml:space="preserve">${text}</w:t></w:r>`;
let id = 0;
const ins = (rsid, text, date) => `<w:ins w:id="${++id}" w:author="Student" w:date="${date}">${run(rsid, text)}</w:ins>`;
const del = (rsid, text, date) =>
  `<w:del w:id="${++id}" w:author="Student" w:date="${date}"><w:r w:rsidR="${rsid}"><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const para = (...runs) => `<w:p>${runs.join("")}</w:p>`;

async function docx(dir, name, settingsXml, body) {
  const zip = new JSZip();
  zip.file("word/document.xml", `<w:document ${W}><w:body>${para(run("00A00001", "Thesis"))}${body}</w:body></w:document>`);
  zip.file("word/settings.xml", settingsXml);
  fs.mkdirSync(dir, { recursive: true });
  const p = path.join(dir, name);
  fs.writeFileSync(p, await zip.generateAsync({ type: "nodebuffer" }));
  return p;
}

const template = await docx(tmp, "template.docx", settings({ rsids: ["00A00001"] }), "");
const s1Body = para(
  ins("00B00001", "Chapter one says hello. ", "2026-06-01T10:00:00Z"),
  ins("00B00001", "It continues here.", "2026-06-01T10:05:00Z")
);
// 第2回: 第1回を承諾 (w:ins を外す) し、ロックをかけ直した (salt が変わる) ものに書き足す
const s2Body =
  para(run("00B00001", "Chapter one says "), del("00B00001", "hello", "2026-06-15T09:00:00Z"), ins("00C00001", "goodbye", "2026-06-15T09:00:10Z"), run("00B00001", ". It continues here.")) +
  para(ins("00C00001", "Chapter two begins.", "2026-06-15T09:30:00Z"));
// 第3回: 第2回を承諾せずに返したものに書き足す。ただし記録をオフにして入力した段落 (00E00001) がある
const s3Body =
  s2Body + para(ins("00D00001", "Chapter three.", "2026-07-01T09:00:00Z")) + para(run("00E00001", "Pasted without tracking."));
// 第4回: 第3回の最後の変更 (7/1) より前の日時の変更が初めて現れる
const s4Body = s3Body + para(ins("00F00001", "Backdated text.", "2026-06-20T09:00:00Z"));

const rs = (...x) => ["00A00001", ...x];
const sprint = (n) => path.join(tmp, `sprint${n}`, "tanaka");
await docx(sprint(1), "thesis.docx", settings({ rsids: rs("00B00001") }), s1Body);
await docx(sprint(2), "thesis.docx", settings({ rsids: rs("00B00001", "00C00001"), salt: "S1" }), s2Body);
await docx(sprint(3), "thesis.docx", settings({ rsids: rs("00B00001", "00C00001", "00D00001", "00E00001"), salt: "S2" }), s3Body);
await docx(sprint(4), "thesis.docx", settings({ rsids: rs("00B00001", "00C00001", "00D00001", "00E00001", "00F00001"), salt: "S2" }), s4Body);

const archive = path.join(tmp, "archive");
function snapshot(...args) {
  const r = spawnSync(process.execPath, [cli, archive, ...args, "--lang", "en"], { encoding: "utf-8" });
  if (r.status !== 0) console.error(r.stderr);
  return r;
}

for (let n = 1; n <= 4; n++) {
  const r = snapshot(path.join(tmp, `sprint${n}`), "--template", template);
  check(`sprint ${n}: runs`, r.status === 0, r.stderr);
  check(`sprint ${n}: archived as #${n}`, r.stderr.includes(`tanaka/thesis: archived submission #${n}`), r.stderr);
}
const again = snapshot(path.join(tmp, "sprint4"));
check("the same file is not archived twice", again.stderr.includes("same as submission #4"), again.stderr);

const dir = path.join(archive, "tanaka", "thesis");
const files = fs.readdirSync(dir).sort();
for (const f of ["thesis-s01.docx", "thesis-s02.docx", "thesis-s03.docx", "thesis-s04.docx", "snapshots.json", "report.json"]) {
  check(`archive has ${f}`, files.includes(f), files.join(", "));
}

const report = JSON.parse(fs.readFileSync(path.join(dir, "report.json"), "utf-8"));
const s = report.snapshots;
const ids = (i) => s[i].evidence.map((e) => e.id).sort().join(",");
check("#1: 2 new changes", s[0].newChanges === 2 && s[0].carriedChanges === 0, JSON.stringify(s[0]));
check("#1: no trace (checked against the template)", ids(0) === "", ids(0));
check("#2: 3 new changes after acceptance", s[1].newChanges === 3 && s[1].carriedChanges === 0, JSON.stringify(s[1]));
check("#2: re-locking by the teacher is not a trace", ids(1) === "", ids(1));
check("#3: changes left from #2 are carried over, not counted again", s[2].newChanges === 1 && s[2].carriedChanges === 3, JSON.stringify(s[2]));
check("#3: text typed without tracking is found", ids(2) === "untrackedText", ids(2));
check("#4: a change dated before #3 is found", ids(3) === "datedBeforePrevious", ids(3));

check("#1 chart has the normal name", files.includes("thesis-s01.svg") && files.includes("thesis-s01-flow.svg"));
check("#3 chart is marked as tampered", files.includes("thesis-s03-tampered.svg") && files.includes("thesis-s03-flow-tampered.svg"));
check("through chart is marked as tampered", files.includes("thesis-through-tampered.svg") && !files.includes("thesis-through.svg"));
const through = fs.readFileSync(path.join(dir, "thesis-through-tampered.svg"), "utf-8");
check("through chart is well-formed", XMLValidator.validate(through) === true);
check("through chart labels each submission", ["#1", "#2", "#3", "#4"].every((l) => through.includes(`>${l}<title>`)));
check("through chart warns per submission", through.includes("#3: ") && through.includes("#4: "));

// 提出物を渡さずに実行すると、図と一覧だけを作り直す
fs.rmSync(path.join(archive, "index.html"));
const redraw = snapshot();
check("redraw without inputs", redraw.status === 0 && fs.existsSync(path.join(archive, "index.html")), redraw.stderr);
const csv = fs.readFileSync(path.join(archive, "summary.csv"), "utf-8");
check("summary.csv has a row per submission", csv.trim().split(/\r?\n/).length === 5, csv);
check("summary.csv links the figures", csv.includes("tanaka/thesis/thesis-s03-tampered.svg"));

// --key-depth: 提出のたびにファイル名が変わっても、学生のフォルダで同じ文書とみなす
await docx(path.join(tmp, "renamed", "tanaka"), "thesis-v5.docx", settings({ rsids: rs("00B00001") }), s4Body + para(ins("00G00001", "More.", "2026-07-10T09:00:00Z")));
const keyed = snapshot(path.join(tmp, "renamed"), "--key-depth", "1");
check("--key-depth groups by the student folder", keyed.stderr.includes("tanaka: archived submission #1"), keyed.stderr);

// KEEP_TMP=1 なら、出力を確かめられるよう一時フォルダを残す
if (process.env.KEEP_TMP) console.log(`kept ${tmp}`);
else fs.rmSync(tmp, { recursive: true, force: true });
if (failed) {
  console.error(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("snapshots: all checks passed");
