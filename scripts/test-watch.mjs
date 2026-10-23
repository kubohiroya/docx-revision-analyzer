// 文書の監視 (src/node/docWatcher.ts) のテスト。
//
// 一時フォルダの .docx を DocWatcher で監視し、次のことを確かめる。
//   - 監視の開始: 最初の版の保管とコメントの記録、「保全済み」の返信の待ち行列
//   - 保存の検知: 新しい版の保管、コメントの出来事 (新しいコメント・解決済み)、同じ内容の保存は無視
//   - 文書が開いている間 (Word の所有者ファイル ~$… がある間) は書き込まず、閉じたら返信を書き込む
//   - 返信の書き込みは新しい版として扱わず、返信のコメントは記録しない
//   - 監視の再開: 監視していない間の変化を記録し、追いつく。移動・改名した文書は同じ記録につなぐ
//
//   node scripts/test-watch.mjs
import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const JSZip = require("jszip");

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-watch-"));
const lib = path.join(tmp, "lib.cjs");
await esbuild.build({ entryPoints: [path.join(root, "src/index.ts")], bundle: true, platform: "node", format: "cjs", outfile: lib, logLevel: "silent" });
const { DocWatcher, isOpenInWord, extractComments, RECEIPT_AUTHOR } = require(lib);

let failed = 0;
function check(name, cond, detail = "") {
  if (cond) console.log(`ok ${name}`);
  else {
    failed++;
    console.error(`FAIL ${name}${detail ? `: ${detail}` : ""}`);
  }
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function until(cond, ms = 4000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await cond()) return true;
    await sleep(50);
  }
  return false;
}

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

/** comments: [{ id, text, done }] (すべて段落 0 の範囲に付ける) */
async function docx(body, comments = [], docId = "{11111111-2222-3333-4444-555555555555}") {
  const zip = new JSZip();
  const starts = comments.map((c) => `<w:commentRangeStart w:id="${c.id}"/>`).join("");
  const ends = comments.map((c) => `<w:commentRangeEnd w:id="${c.id}"/><w:r><w:commentReference w:id="${c.id}"/></w:r>`).join("");
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`
  );
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
  zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8"?><w:settings ${NS}><w:trackRevisions/><w15:docId w15:val="${docId}"/></w:settings>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="00000001">${starts}${run(body)}${ends}</w:p><w:sectPr/></w:body></w:document>`);
  if (comments.length) {
    zip.file(
      "word/comments.xml",
      `<?xml version="1.0" encoding="UTF-8"?><w:comments ${NS}>` +
        comments.map((c) => `<w:comment w:id="${c.id}" w:author="Teacher" w:date="2026-06-01T10:00:00Z"><w:p w14:paraId="0000010${c.id}">${run(c.text)}</w:p></w:comment>`).join("") +
        `</w:comments>`
    );
    zip.file(
      "word/commentsExtended.xml",
      `<?xml version="1.0" encoding="UTF-8"?><w15:commentsEx ${NS}>` +
        comments.map((c) => `<w15:commentEx w15:paraId="0000010${c.id}" w15:done="${c.done ? 1 : 0}"/>`).join("") +
        `</w15:commentsEx>`
    );
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

const docDir = path.join(tmp, "docs");
const records = path.join(tmp, "records");
fs.mkdirSync(docDir);
const doc = path.join(docDir, "report.docx");
fs.writeFileSync(doc, await docx("first draft", [{ id: 0, text: "Cite a source." }]));

// Word の所有者ファイルによる判定
{
  check("isOpenInWord: closed", !(await isOpenInWord(doc)));
  fs.writeFileSync(path.join(docDir, "~$report.docx"), "owner");
  check("isOpenInWord: owner file (short name)", await isOpenInWord(doc));
  fs.rmSync(path.join(docDir, "~$report.docx"));
  const long = path.join(docDir, "graduation-thesis.docx");
  fs.writeFileSync(path.join(docDir, "~$aduation-thesis.docx"), "owner");
  check("isOpenInWord: owner file (first two characters replaced)", await isOpenInWord(long));
  check("isOpenInWord: other document's owner file ignored", !(await isOpenInWord(doc)));
  fs.rmSync(path.join(docDir, "~$aduation-thesis.docx"));
}

let open = true;
const updates = [];
const receipts = [];
const errors = [];
const options = {
  path: doc,
  recordsRoot: records,
  receiptText: (seq) => `Collected #${seq}`,
  onUpdate: (u) => updates.push(u),
  onReceipts: (r) => receipts.push(r),
  onError: (e) => errors.push(e),
  debounceMs: 150,
  closePollMs: 100,
  statPollMs: 200,
  isOpen: async () => open,
};

// 監視の開始
const w = new DocWatcher(options);
const first = await w.start();
const dir = w.archive.dir;
check("start: records folder", fs.existsSync(path.join(dir, "watch.json")) && dir.startsWith(records));
check("start: first version archived", first.version?.n === 1 && fs.existsSync(path.join(dir, first.version.file)));
check("start: comment recorded as seen", first.events.length === 1 && first.events[0].type === "seen" && first.events[0].seq === 1);
check("start: receipt queued", w.archive.receipts().pending.length === 1);
check("start: document id recorded", w.archive.info.docId === "{11111111-2222-3333-4444-555555555555}");

// 保存の検知 (開いている間): 新しいコメント、最初のコメントは解決済み
fs.writeFileSync(doc, await docx("second draft", [{ id: 0, text: "Cite a source.", done: true }, { id: 1, text: "Shorten this." }]));
check("save: update received", await until(() => updates.length === 1), String(updates.length));
const u1 = updates[0];
check("save: new version archived", u1?.version?.n === 2);
check("save: events (resolved, seen)", JSON.stringify(u1?.events.map((e) => e.type).sort()) === JSON.stringify(["resolved", "seen"]));
check("save: two receipts queued", w.archive.receipts().pending.length === 2);
await sleep(400);
check("open: nothing written while open", receipts.length === 0 && !(await extractComments(fs.readFileSync(doc))).some((c) => c.author === RECEIPT_AUTHOR));

// 同じ内容の保存は無視
fs.writeFileSync(doc, fs.readFileSync(doc));
await sleep(500);
check("save: same content ignored", updates.length === 1, String(updates.length));

// 閉じたら返信を書き込む
open = false;
check("close: receipts written", await until(() => receipts.length === 1));
const written = await extractComments(fs.readFileSync(doc));
const replies = written.filter((c) => c.author === RECEIPT_AUTHOR);
check("close: two replies in the document", replies.length === 2 && replies.some((c) => c.text === "Collected #1") && replies.some((c) => c.text === "Collected #3"), JSON.stringify(replies.map((c) => c.text)));
check("close: queue emptied", w.archive.receipts().pending.length === 0 && w.archive.receipts().written.length === 2);
await sleep(600);
check("close: tool write is not a new version", w.archive.versions().length === 2 && updates.length === 1, `${w.archive.versions().length} / ${updates.length}`);
check("close: receipt events recorded", w.archive.readEvents().filter((e) => e.type === "receipt").length === 2);
check("close: replies not recorded as comments", w.comments().length === 2 && w.comments().every((c) => c.receiptAt));
check("close: no leftover temp file", !fs.readdirSync(docDir).some((f) => f.includes(".~dra-")));
w.stop();
check("stop: period closed", !!w.archive.info.periods[0].end);
check("no watcher errors", errors.length === 0, errors.map(String).join("; "));

// 監視していない間の変更: 再開時に追いつき、記録する
fs.writeFileSync(doc, await docx("edited while not watched", [{ id: 0, text: "Cite a source.", done: true }]));
const w2 = new DocWatcher({ ...options, isOpen: async () => true });
const caught = await w2.start();
check("resume: same records folder", w2.archive.dir === dir);
check("resume: catch-up version and gap", caught.kind === "catchUp" && caught.gap && caught.version?.n === 3 && w2.archive.info.gaps.length === 1);
check("resume: removed comment recorded", caught.events.some((e) => e.type === "removed"));
check("resume: removed comment kept in the record", w2.comments().find((c) => !c.present)?.comment.text === "Shorten this.");
w2.stop();

// 移動・改名した文書: 識別子と内容が同じなら同じ記録につなぐ
const moved = path.join(docDir, "renamed.docx");
fs.renameSync(doc, moved);
const w3 = new DocWatcher({ ...options, path: moved, isOpen: async () => true });
const m = await w3.start();
check("move: same records folder", w3.archive.dir === dir && w3.archive.info.paths.length === 2 && !m.gap);
w3.stop();

// 書き込めない読み取り元 (OneDrive の文書の問い合わせの代わり): 版とコメントは記録し、返信は待ち行列に入れない
{
  let bytes = await docx("cloud draft", [{ id: 0, text: "Cloud comment." }], "{99999999-0000-0000-0000-000000000000}");
  let notify;
  const cloudUpdates = [];
  const cw = new DocWatcher({
    ...options,
    path: undefined,
    source: {
      location: "https://contoso.sharepoint.com/sites/a/Shared%20Documents/cloud.docx",
      name: "cloud.docx",
      canWrite: false,
      read: async () => ({ bytes, mtime: new Date() }),
      subscribe: (cb) => {
        notify = cb;
        return () => (notify = undefined);
      },
    },
    onUpdate: (u) => cloudUpdates.push(u),
  });
  const c0 = await cw.start();
  check("cloud: records folder keyed by the URL", c0.version?.n === 1 && cw.archive.dir !== dir && cw.archive.info.path.startsWith("https://"));
  check("cloud: comment recorded, no receipt queued", c0.events[0]?.type === "seen" && cw.archive.receipts().pending.length === 0);
  bytes = await docx("cloud draft 2", [{ id: 0, text: "Cloud comment.", done: true }], "{99999999-0000-0000-0000-000000000000}");
  notify();
  check("cloud: change picked up", await until(() => cloudUpdates.length === 1) && cloudUpdates[0].version?.n === 2 && cloudUpdates[0].events[0]?.type === "resolved");
  cw.stop();
}

if (failed) {
  console.error(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("watch: all checks passed");
fs.rmSync(tmp, { recursive: true, force: true });
