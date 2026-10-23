// Word のコメントの読み取り・記録・返信の書き込み (src/lib/docxComments.ts) のテスト。
//
// Word 2019 以降と同じ形のコメント (comments.xml・commentsExtended.xml・commentsIds.xml・commentsExtensible.xml) を
// 持つ文書を JSZip で作り、次のことを確かめる。
//   - 本文・作成者・UTC の日時・durableId・解決済み・返信の親子・範囲 (段落・引用・前後・重なる変更) を読めること
//   - 出来事の列 (初めて見た・本文の変更・解決・削除・再出現) と、その再生
//   - 返信の書き込み: 親子の関係が付き、本文の文字・変更履歴・既存のコメント・整合性の検査結果が変わらないこと
//   - commentsExtended.xml などの無い古い形の文書にも返信を書けること
//
//   node scripts/test-comments.mjs
import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const { XMLValidator } = require("fast-xml-parser");

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-comments-"));
const lib = path.join(tmp, "core.cjs");
await esbuild.build({ entryPoints: [path.join(root, "src/core.ts")], bundle: true, platform: "node", format: "cjs", outfile: lib, logLevel: "silent" });
const {
  extractComments,
  diffComments,
  replayCommentEvents,
  addCommentReplies,
  RECEIPT_AUTHOR,
  extractRevisions,
  parseDocxLayout,
  finalDocumentText,
  checkIntegrity,
  checkTamperEvidence,
  hasTamperEvidence,
} = require(lib);

let failed = 0;
function check(name, cond, detail = "") {
  if (cond) console.log(`ok ${name}`);
  else {
    failed++;
    console.error(`FAIL ${name}${detail ? `: ${detail}` : ""}`);
  }
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS = `xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const T1 = "2026-06-01T10:00:00Z";
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const ref = (id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`;

/** 文書: コメント 0 (範囲に変更の挿入を含む) と、その返信 1 (解決済み)、範囲の印の無いコメント 2 */
function documentXml() {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    `<w:p w14:paraId="1A2B3C4D">${run("Before ")}<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>${run("quoted ")}` +
    `<w:ins w:id="10" w:author="Writer" w:date="${T1}">${run("inserted")}</w:ins>` +
    `<w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="1"/>${ref(0)}${ref(1)}${run(" after")}</w:p>` +
    `<w:p w14:paraId="2B3C4D5E">${run("Second paragraph")}${ref(2)}</w:p>` +
    `<w:sectPr/></w:body></w:document>`
  );
}

const comment = (id, author, date, text, paraId) =>
  `<w:comment w:id="${id}" w:author="${author}" w:date="${date}" w:initials="${author[0]}"><w:p w14:paraId="${paraId}"><w:r><w:annotationRef/></w:r>${run(text)}</w:p></w:comment>`;

function commentsXml(comments) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${NS}>${comments.join("")}</w:comments>`;
}

const MODERN = {
  "word/commentsExtended.xml":
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">` +
    `<w15:commentEx w15:paraId="00000A01" w15:done="0"/><w15:commentEx w15:paraId="00000A02" w15:paraIdParent="00000A01" w15:done="1"/>` +
    `<w15:commentEx w15:paraId="00000A03" w15:done="0"/></w15:commentsEx>`,
  "word/commentsIds.xml":
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">` +
    `<w16cid:commentId w16cid:paraId="00000A01" w16cid:durableId="11111111"/><w16cid:commentId w16cid:paraId="00000A02" w16cid:durableId="22222222"/>` +
    `<w16cid:commentId w16cid:paraId="00000A03" w16cid:durableId="33333333"/></w16cid:commentsIds>`,
  "word/commentsExtensible.xml":
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">` +
    `<w16cex:commentExtensible w16cex:durableId="11111111" w16cex:dateUtc="2026-06-01T01:00:00Z"/>` +
    `<w16cex:commentExtensible w16cex:durableId="22222222" w16cex:dateUtc="2026-06-01T02:00:00Z"/></w16cex:commentsExtensible>`,
};

const BASE_COMMENTS = [
  comment(0, "Teacher", "2026-06-01T10:00:00Z", "Please cite a source.", "00000A01"),
  comment(1, "Writer", "2026-06-01T11:00:00Z", "Added.", "00000A02"),
  comment(2, "Writer", "2026-06-01T12:00:00Z", "Check the wording.", "00000A03"),
];

async function docx({ comments = BASE_COMMENTS, modern = true } = {}) {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`
  );
  zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NS}><w:trackRevisions/></w:settings>`);
  zip.file("word/document.xml", documentXml());
  zip.file("word/comments.xml", commentsXml(comments));
  if (modern) for (const [name, xml] of Object.entries(MODERN)) zip.file(name, xml);
  return zip.generateAsync({ type: "uint8array" });
}

// ---------------------------------------------------------------------------
// 読み取り
{
  const cs = await extractComments(await docx());
  const [a, b, c] = cs;
  check("extract: three comments", cs.length === 3, String(cs.length));
  check("extract: durable key", a.key === "d:11111111" && b.key === "d:22222222");
  check("extract: author, text, utc date", a.author === "Teacher" && a.text === "Please cite a source." && a.dateUtc === "2026-06-01T01:00:00Z");
  check("extract: reply parent and resolved", b.parentKey === a.key && b.done === true && a.done === false);
  check("extract: anchor paragraph and quote", a.anchor?.paraId === "1A2B3C4D" && a.anchor.quote === "quoted inserted", JSON.stringify(a.anchor));
  check("extract: anchor context", a.anchor.prefix.endsWith("Before ") && a.anchor.suffix.startsWith(" after"));
  check("extract: anchor overlapping revision", a.anchor.revisions.length === 1 && a.anchor.revisions[0].startsWith(`ins|Writer|${T1}|`));
  check("extract: point anchor without range", c.anchor?.paraId === "2B3C4D5E" && c.anchor.quote === "");

  const old = await extractComments(await docx({ modern: false }));
  check("extract (old format): hashed keys", old.every((x) => x.key.startsWith("h:")) && new Set(old.map((x) => x.key)).size === 3);
  check("extract (old format): no parent, not resolved", old.every((x) => !x.parentKey && !x.done));
}

// ---------------------------------------------------------------------------
// 出来事の列
{
  const t0 = new Date("2026-06-02T00:00:00Z");
  const events = [];
  const record = (list) => list.forEach((e) => events.push({ ...e, seq: events.length + 1 }));
  record(diffComments(new Map(), await extractComments(await docx()), t0, "sha-1"));
  check("events: all comments seen", events.length === 3 && events.every((e) => e.type === "seen"));

  // 本文の変更・解決の取り消し・削除
  const changed = [
    comment(0, "Teacher", "2026-06-01T10:00:00Z", "Please cite two sources.", "00000A01"),
    comment(1, "Writer", "2026-06-01T11:00:00Z", "Added.", "00000A02"),
  ];
  const modern = await docx({ comments: changed });
  const zip = await JSZip.loadAsync(modern);
  zip.file("word/commentsExtended.xml", MODERN["word/commentsExtended.xml"].replace('w15:done="1"', 'w15:done="0"'));
  record(diffComments(replayCommentEvents(events), await extractComments(await zip.generateAsync({ type: "uint8array" })), t0, "sha-2"));
  const types = events.slice(3).map((e) => e.type).sort();
  check("events: edited, reopened, removed", JSON.stringify(types) === JSON.stringify(["edited", "removed", "reopened"]), JSON.stringify(types));
  const state = replayCommentEvents(events);
  check("replay: edited text and record number kept", state.get("d:11111111").comment.text === "Please cite two sources." && state.get("d:11111111").seq === 1);
  check("replay: removed comment kept in the record", state.get("d:33333333").present === false && state.get("d:33333333").comment.text === "Check the wording.");

  record(diffComments(state, await extractComments(await docx()), t0, "sha-3"));
  const back = events.slice(6).map((e) => e.type).sort();
  check("events: reappeared after removal", back.includes("reappeared"), JSON.stringify(back));
}

// ---------------------------------------------------------------------------
// 返信の書き込み
{
  const input = await docx();
  const before = await extractComments(input);
  const r = await addCommentReplies(input, [
    { key: "d:11111111", text: "Collected <#1> & preserved" },
    { key: "d:33333333", text: "Collected #3" },
    { key: "d:99999999", text: "missing" },
  ], { now: new Date("2026-06-03T00:00:00Z") });
  check("replies: two written, one skipped", r.written.length === 2 && r.skipped.length === 1 && r.skipped[0].reason === "missing");
  const after = await extractComments(r.bytes);
  const receipts = after.filter((c) => c.author === RECEIPT_AUTHOR);
  check("replies: two receipts", receipts.length === 2, String(receipts.length));
  check("replies: threaded under the parent", receipts.some((c) => c.parentKey === "d:11111111" && c.text === "Collected <#1> & preserved") && receipts.some((c) => c.parentKey === "d:33333333"));
  check("replies: durable ids and utc dates", receipts.every((c) => c.durableId && c.dateUtc === "2026-06-03T00:00:00Z"));
  check("replies: receipt anchored at the parent's range", receipts.find((c) => c.parentKey === "d:11111111").anchor.quote === "quoted inserted");
  const kept = after.filter((c) => c.author !== RECEIPT_AUTHOR);
  check("replies: existing comments unchanged", JSON.stringify(kept) === JSON.stringify(before));
  check("replies: receipts are not recorded", diffComments(replayCommentEvents([]), after, new Date()).length === 3);

  const z = await JSZip.loadAsync(r.bytes);
  for (const name of ["word/document.xml", "word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml"]) {
    check(`replies: ${name} is well-formed`, XMLValidator.validate(await z.file(name).async("string")) === true);
  }
  const textOf = async (b) => finalDocumentText(await parseDocxLayout(b));
  check("replies: document text unchanged", (await textOf(input)) === (await textOf(r.bytes)));
  const revs = async (b) => JSON.stringify((await extractRevisions(b)).events.map((e) => [e.type, e.author, e.date, e.text]));
  check("replies: tracked changes unchanged", (await revs(input)) === (await revs(r.bytes)));
  const integ = async (b) => JSON.stringify((await checkIntegrity(b)).items.map((i) => [i.id ?? i.kind, i.observed]));
  check("replies: integrity check unchanged", (await integ(input)) === (await integ(r.bytes)), `${await integ(input)} / ${await integ(r.bytes)}`);
  check("replies: no tamper evidence from the receipts", hasTamperEvidence(await checkTamperEvidence(input)) === hasTamperEvidence(await checkTamperEvidence(r.bytes)));

  // 古い形の文書: commentsExtended.xml を作り、関係と型を登録する
  const old = await docx({ modern: false });
  const oldKeys = (await extractComments(old)).map((c) => c.key);
  const r2 = await addCommentReplies(old, [{ key: oldKeys[0], text: "Collected" }]);
  const z2 = await JSZip.loadAsync(r2.bytes);
  check("old format: commentsExtended.xml created", !!z2.file("word/commentsExtended.xml"));
  check("old format: relationship and content type", (await z2.file("word/_rels/document.xml.rels").async("string")).includes("commentsExtended.xml") && (await z2.file("[Content_Types].xml").async("string")).includes("/word/commentsExtended.xml"));
  const after2 = await extractComments(r2.bytes);
  check("old format: reply threaded", after2.some((c) => c.author === RECEIPT_AUTHOR && c.parentKey === oldKeys[0]));
}

if (failed) {
  console.error(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("comments: all checks passed");
fs.rmSync(tmp, { recursive: true, force: true });
