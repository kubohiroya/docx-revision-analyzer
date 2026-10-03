// 指定時刻の版の復元 (src/lib/versions.ts, src/cli/versions.ts) のテスト。
//
// 変更の種類ごとに日時をずらした文書を JSZip で作り、docx-revision-versions を実際に動かして、各時刻の版の本文・
// 書式・出力ファイル・versions.json を確かめる。
//
//   T1 (10:00) 語句の挿入
//   T2 (10:10) 挿入 (後で T4 に削除される) と、段落の分割 (段落記号の挿入)
//   T3 (10:20) 語句の削除、段落の削除 (段落記号ごと)、太字にする書式の変更
//   T4 (10:30) T2 の挿入の削除、表の行の挿入
//   T5 (10:40) 段落の挿入
//   T6 (翌日)  段落の移動 (末尾から先頭へ)
//   日時の無い挿入 (最初から反映済みとして扱う)
//
//   node scripts/test-versions.mjs            (KEEP_TMP=1 で出力を残す)
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

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-versions-"));
const cli = path.join(tmp, "versions.cjs");
await esbuild.build({
  entryPoints: [path.join(root, "src/cli/versions.ts")],
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
const T = {
  1: "2026-06-01T10:00:00Z",
  2: "2026-06-01T10:10:00Z",
  3: "2026-06-01T10:20:00Z",
  4: "2026-06-01T10:30:00Z",
  5: "2026-06-01T10:40:00Z",
  6: "2026-06-02T09:00:00Z",
};
let id = 0;
const rev = (date) => `w:id="${++id}" w:author="Writer"${date ? ` w:date="${date}"` : ""}`;
const run = (text, rPr = "") => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const delRun = (text) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`;
const ins = (date, text) => `<w:ins ${rev(date)}>${run(text)}</w:ins>`;
const del = (date, text) => `<w:del ${rev(date)}>${delRun(text)}</w:del>`;
/** 段落記号の挿入・削除・移動の印 */
const mark = (kind, date) => `<w:pPr><w:rPr><w:${kind} ${rev(date)}/></w:rPr></w:pPr>`;
const p = (...parts) => `<w:p>${parts.join("")}</w:p>`;

const body = [
  // 移動先 (T6)
  p(
    mark("moveTo", T[6]),
    `<w:moveToRangeStart w:id="900" w:name="move1" ${rev(T[6]).replace(/w:id="\d+" /, "")}/>`,
    `<w:moveTo ${rev(T[6])}>${run("Moved")}</w:moveTo>`,
    `<w:moveToRangeEnd w:id="900"/>`
  ),
  // 語句の挿入・削除、挿入の後の削除、日時の無い挿入
  p(
    ins(undefined, "Undated "),
    run("Alpha "),
    ins(T[1], "beta "),
    run("gamma"),
    `<w:ins ${rev(T[2])}><w:del ${rev(T[4])}>${delRun("temp")}</w:del></w:ins>`,
    del(T[3], "delta")
  ),
  // 段落の分割 (T2 に段落記号を挿入)
  p(mark("ins", T[2]), run("Left")),
  p(run("Right")),
  // 段落の削除 (T3)
  p(mark("del", T[3]), del(T[3], "Gone")),
  p(run("Stays")),
  // 段落の挿入 (T5)
  p(mark("ins", T[5]), ins(T[5], "New para")),
  p(run("Last")),
  // 表の行の挿入 (T4)
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
    `<w:tr><w:tc>${p(run("Cell A"))}</w:tc></w:tr>` +
    `<w:tr><w:trPr><w:ins ${rev(T[4])}/></w:trPr><w:tc>${p(mark("ins", T[4]), ins(T[4], "Cell B"))}</w:tc></w:tr></w:tbl>`,
  // 書式の変更 (T3 に太字にした)
  p(run("Bold", `<w:rPr><w:b/><w:rPrChange ${rev(T[3])}><w:rPr/></w:rPrChange></w:rPr>`)),
  // 移動元 (T6)
  p(
    mark("moveFrom", T[6]),
    `<w:moveFromRangeStart w:id="901" w:name="move1" ${rev(T[6]).replace(/w:id="\d+" /, "")}/>`,
    `<w:moveFrom ${rev(T[6])}>${delRun("Moved")}</w:moveFrom>`,
    `<w:moveFromRangeEnd w:id="901"/>`
  ),
  p(run("End")),
].join("");

const zip = new JSZip();
zip.file(
  "[Content_Types].xml",
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
);
zip.file(
  "_rels/.rels",
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
);
zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
const input = path.join(tmp, "essay.docx");
fs.writeFileSync(input, await zip.generateAsync({ type: "nodebuffer" }));

function versions(...args) {
  const r = versionsQuiet(...args);
  if (r.status !== 0) console.error(r.stderr);
  return r;
}
/** 失敗を確かめるとき用 (エラーを表示しない) */
function versionsQuiet(...args) {
  return spawnSync(process.execPath, [cli, input, ...args, "--lang", "en"], { encoding: "utf-8" });
}
const readJson = (dir) => JSON.parse(fs.readFileSync(path.join(dir, "essay.versions.json"), "utf-8"));
const lines = (dir, file) => fs.readFileSync(path.join(dir, file), "utf-8").trimEnd().split("\n");

// --at: 各時刻の版の本文
const atDir = path.join(tmp, "at");
const atTimes = [
  ["before T1", "2026-06-01T09:59:59Z"],
  ["T2", T[2]],
  ["T3", T[3]],
  ["T4", T[4]],
  ["T5", T[5]],
  ["T6", T[6]],
];
const r = versions("-o", atDir, "-f", "docx,txt", ...atTimes.flatMap(([, t]) => ["--at", t]));
check("--at runs", r.status === 0, r.stderr);
check("undated changes are reported", r.stderr.includes("1 tracked change(s) have no date"), r.stderr);
const manifest = readJson(atDir);
check("one version per --at", manifest.versions.length === atTimes.length, JSON.stringify(manifest.versions));
check("versions are labelled 'at'", manifest.versions.every((v) => v.label === "at"));
const expected = {
  "before T1": ["Undated Alpha gammadelta", "LeftRight", "Gone", "Stays", "Last", "Cell A", "Bold", "Moved", "End"],
  T2: ["Undated Alpha beta gammatempdelta", "Left", "Right", "Gone", "Stays", "Last", "Cell A", "Bold", "Moved", "End"],
  T3: ["Undated Alpha beta gammatemp", "Left", "Right", "Stays", "Last", "Cell A", "Bold", "Moved", "End"],
  T4: ["Undated Alpha beta gamma", "Left", "Right", "Stays", "Last", "Cell A", "Cell B", "Bold", "Moved", "End"],
  T5: ["Undated Alpha beta gamma", "Left", "Right", "Stays", "New para", "Last", "Cell A", "Cell B", "Bold", "Moved", "End"],
  T6: ["Moved", "Undated Alpha beta gamma", "Left", "Right", "Stays", "New para", "Last", "Cell A", "Cell B", "Bold", "End"],
};
for (const [i, [label]] of atTimes.entries()) {
  const v = manifest.versions[i];
  const txt = v.files.find((f) => f.endsWith(".txt"));
  const got = lines(atDir, txt);
  check(`text at ${label}`, JSON.stringify(got) === JSON.stringify(expected[label]), JSON.stringify(got));
  const xml = await (await JSZip.loadAsync(fs.readFileSync(path.join(atDir, v.files.find((f) => f.endsWith(".docx")))))).file("word/document.xml").async("string");
  check(`docx at ${label} is well-formed`, XMLValidator.validate(xml) === true);
  check(
    `docx at ${label} keeps no tracked changes`,
    !/<w:(ins|del|moveFrom|moveTo|delText|rPrChange|moveFromRangeStart|moveToRangeStart)[ >/]/.test(xml),
    xml.match(/<w:(ins|del|moveFrom|moveTo|delText|rPrChange)[^>]*>/)?.[0]
  );
  const bold = /<w:rPr><w:b\/><\/w:rPr><w:t[^>]*>Bold</.test(xml);
  check(`bold at ${label} follows the format change`, bold === !["before T1", "T2"].includes(label), xml.match(/<w:r>.{0,80}Bold</)?.[0]);
}
check("version file names carry the time", manifest.versions[1].files.includes(`essay-${stamp(T[2])}.docx`), manifest.versions[1].files.join(","));

// 既定 (--sessions): 最初の変更の直前、1日目の区間の終わり (T5)、2日目 (T6)
const sDir = path.join(tmp, "sessions");
const s = versions("-o", sDir);
check("--sessions (default) runs", s.status === 0, s.stderr);
const sm = readJson(sDir);
check(
  "--sessions makes the initial version and one per session",
  JSON.stringify(sm.versions.map((v) => [v.label, v.at])) ===
    JSON.stringify([["initial", "2026-06-01T09:59:59.999Z"], ["session", "2026-06-01T10:40:00.000Z"], ["session", T[6].replace("Z", ".000Z")]]),
  JSON.stringify(sm.versions.map((v) => [v.label, v.at]))
);
check("--sessions writes .docx only by default", sm.versions.every((v) => v.files.length === 1 && v.files[0].endsWith(".docx")));
check("changes are counted up to each version", JSON.stringify(sm.versions.map((v) => v.changes)) === JSON.stringify([0, sm.totalChanges - 4, sm.totalChanges]), JSON.stringify(sm.versions.map((v) => v.changes)));

// --every: 変更の無い区切りは飛ばす
const eDir = path.join(tmp, "every");
const e = versions("-o", eDir, "--every", "15m", "-f", "txt");
check("--every runs", e.status === 0, e.stderr);
const em = readJson(eDir);
check(
  "--every skips intervals with no changes",
  JSON.stringify(em.versions.map((v) => v.at.slice(11, 16))) === JSON.stringify(["09:59", "10:15", "10:30", "10:45", "09:00"]),
  JSON.stringify(em.versions.map((v) => v.at))
);
check("--every with a bad interval fails", versionsQuiet("-o", eDir, "--every", "soon").status !== 0);
check("--format with a bad value fails", versionsQuiet("-o", eDir, "-f", "pdf").status !== 0);

/** ISO 時刻を、ファイル名に使うローカル時刻 (YYYYMMDD-HHMMSS) にする */
function stamp(iso) {
  const d = new Date(iso);
  const z = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}${z(d.getSeconds())}`;
}

// KEEP_TMP=1 なら、出力を確かめられるよう一時フォルダを残す
if (process.env.KEEP_TMP) console.log(`kept ${tmp}`);
else fs.rmSync(tmp, { recursive: true, force: true });
if (failed) {
  console.error(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("versions: all checks passed");
