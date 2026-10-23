// スモークテスト (文書の監視): fixtures の文書にコメントを付けた2つの版を一時フォルダに作り、アプリを起動して
// 1つ目を開いて監視を始め、2つ目で置き換える (保存の代わり)。次のことを確かめて、画面を PNG に保存する。
//   - 記録のフォルダに、保存ごとの版 (2つ) とコメントの記録が残ること
//   - 文書が閉じられている (Word の所有者ファイルが無い) ので、コメントに「保全済み」の返信が書き込まれること
//   - URL から開いた文書 (Graph の代わりに DRA_GRAPH_MOCK のファイル) も、問い合わせで新しい版を拾い、返信は書かないこと
//   - LibreOffice があれば (DRA_SMOKE_SAVED、macOS は既定の場所にあれば自動)、最新の版の保存時のページを右の画面に表示し、
//     その PDF と pages.json が記録のフォルダに残ること (issue #59)
//   npm run smoke:watch
import { spawnSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const dir = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const electron = require("electron");
const JSZip = createRequire(path.join(dir, "../package.json"))("jszip");

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-smoke-watch-"));
const out = path.resolve(process.env.DRA_SMOKE_OUT ?? path.join(dir, "smoke-out"));
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const saved = process.env.DRA_SMOKE_SAVED ?? (fs.existsSync("/Applications/LibreOffice.app/Contents/MacOS/soffice") ? "1" : "");

/** fixtures/flow-demo.docx の最初の段落にコメントを付ける */
async function withComments(texts) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(dir, "../fixtures/flow-demo.docx")));
  let doc = await zip.file("word/document.xml").async("string");
  const starts = texts.map((_, i) => `<w:commentRangeStart w:id="${i}"/>`).join("");
  const ends = texts.map((_, i) => `<w:commentRangeEnd w:id="${i}"/><w:r><w:commentReference w:id="${i}"/></w:r>`).join("");
  doc = doc.replace(/(<w:p\b[^>]*>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)([\s\S]*?)(<\/w:p>)/, `$1${starts}$2${ends}$3`);
  zip.file("word/document.xml", doc);
  zip.file(
    "word/comments.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:comments ${W}>` +
      texts.map((t, i) => `<w:comment w:id="${i}" w:author="Teacher" w:date="2026-06-01T10:00:00Z"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:comment>`).join("") +
      `</w:comments>`
  );
  let rels = (await zip.file("word/_rels/document.xml.rels")?.async("string")) ??
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  rels = rels.replace("</Relationships>", `<Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", rels);
  return zip.generateAsync({ type: "nodebuffer" });
}

const doc = path.join(tmp, "essay.docx");
const next = path.join(tmp, "essay-next.docx");
let failed = 0;
const check = (name, cond, detail = "") => {
  console.log(`${cond ? "ok" : "FAIL"} ${name}${!cond && detail ? `: ${detail}` : ""}`);
  if (!cond) failed++;
};
const readRecords = (userData) => {
  const records = path.join(userData, "watch");
  const folders = fs.existsSync(records) ? fs.readdirSync(records) : [];
  const rec = path.join(records, folders[0] ?? "");
  const read = (f) => (fs.existsSync(path.join(rec, f)) ? fs.readFileSync(path.join(rec, f), "utf-8") : "");
  const manifest = read("snapshots.json") ? JSON.parse(read("snapshots.json")) : { snapshots: [] };
  const events = read("comments.jsonl").trim() ? read("comments.jsonl").trim().split("\n").map((l) => JSON.parse(l)) : [];
  const pages = read("pages.json") ? JSON.parse(read("pages.json")) : { items: {} };
  const pdfs = folders.length ? fs.readdirSync(rec).filter((f) => f.endsWith(".pdf")) : [];
  return { folders, manifest, events, pages, pdfs, info: read("watch.json") ? JSON.parse(read("watch.json")) : {} };
};
/** 保存時のページ: 最新の版の PDF が記録のフォルダにあること */
const checkSavedPages = (label, { manifest, pages, pdfs }) => {
  if (!saved) return console.log(`skip ${label}: saved pages (LibreOffice was not found; set DRA_SMOKE_SAVED=1 to force)`);
  const done = Object.entries(pages.items).filter(([, v]) => v.state === "done");
  const latest = manifest.snapshots[manifest.snapshots.length - 1]?.file?.replace(/\.docx$/i, ".pdf");
  check(`${label}: saved pages of the latest version kept with the records`, done.some(([k]) => k === latest) && pdfs.includes(latest), JSON.stringify({ latest, pages, pdfs }));
};
const replies = async (file) =>
  ((await (await JSZip.loadAsync(fs.readFileSync(file))).file("word/comments.xml").async("string")).match(/w:author="Docx Revision Analyzer"/g) ?? []).length;

// 1. ローカルの文書
{
  fs.writeFileSync(doc, await withComments(["Cite a source."]));
  fs.writeFileSync(next, await withComments(["Cite a source.", "Shorten this paragraph."]));
  const userData = path.join(tmp, "user-data-local");
  const r = spawnSync(electron, [dir], {
    stdio: "inherit",
    env: { ...process.env, DRA_SMOKE_FILE: doc, DRA_SMOKE_WATCH_NEXT: next, DRA_SMOKE_OUT: out, DRA_USER_DATA: userData, DRA_SMOKE_SAVED: saved },
    timeout: 180000,
  });
  check("local: app exited normally", r.status === 0, String(r.status));
  const records = readRecords(userData);
  const { folders, manifest, events } = records;
  check("local: records folder for the document", folders.length === 1, JSON.stringify(folders));
  check("local: two versions archived", manifest.snapshots.length === 2, String(manifest.snapshots.length));
  check("local: both comments recorded", events.filter((e) => e.type === "seen").length === 2, JSON.stringify(events.map((e) => e.type)));
  check("local: receipts written", events.filter((e) => e.type === "receipt").length === 2, JSON.stringify(events.map((e) => e.type)));
  // 1つ目の版への返信は、監視を始めた直後 (文書は閉じている) に書き込まれ、2つ目の版での置き換えで消える。
  // 2つ目の版で新しく見たコメントへの返信が、置き換えた後の文書にあること
  check("local: reply in the document after the save", (await replies(doc)) >= 1);
  checkSavedPages("local", records);
}

// 2. URL から開いた文書 (DRA_GRAPH_MOCK のファイルを OneDrive の文書の代わりにする)
{
  const mock = path.join(tmp, "cloud.docx");
  fs.writeFileSync(mock, await withComments(["Cite a source."]));
  const userData = path.join(tmp, "user-data-cloud");
  const r = spawnSync(electron, [dir], {
    stdio: ["ignore", "pipe", "inherit"],
    encoding: "utf-8",
    env: {
      ...process.env,
      DRA_SMOKE_URL: "https://contoso.sharepoint.com/sites/class/Shared%20Documents/cloud.docx",
      DRA_GRAPH_MOCK: mock,
      DRA_SMOKE_WATCH_NEXT: next,
      DRA_WATCH_CLOUD_POLL_MS: "300",
      DRA_SMOKE_OUT: path.join(out, "cloud"),
      DRA_USER_DATA: userData,
      DRA_SMOKE_SAVED: saved,
    },
    timeout: 180000,
  });
  process.stdout.write(r.stdout ?? "");
  check("cloud: app exited normally", r.status === 0, String(r.status));
  // スモークテストの click の手順は、最初の注釈 (コメントの印) をクリックする: 元の文書の該当箇所を開く
  check("cloud: comment mark opens the source document", /smoke: would open https:\/\/contoso\.sharepoint\.com\//.test(r.stdout ?? ""));
  const records = readRecords(userData);
  const { folders, manifest, events, info } = records;
  check("cloud: records folder keyed by the URL", folders.length === 1 && /^https:/.test(info.path ?? ""), JSON.stringify(folders));
  check("cloud: new version picked up by polling", manifest.snapshots.length === 2, String(manifest.snapshots.length));
  check("cloud: both comments recorded", events.filter((e) => e.type === "seen").length === 2, JSON.stringify(events.map((e) => e.type)));
  check("cloud: no replies written", events.every((e) => e.type !== "receipt") && (await replies(mock)) === 0);
  // OneDrive の文書の最新の版は Graph で PDF にするが、DRA_GRAPH_MOCK では使えないので LibreOffice で作る
  checkSavedPages("cloud", records);
}

fs.rmSync(tmp, { recursive: true, force: true });
process.exit(failed ? 1 : 0);
