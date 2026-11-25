import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "dra-lineage-comments-"));
const require = createRequire(import.meta.url);
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const IDS = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rootText = "#出典確認 元の表を比較した。 https://example.org/report.pdf#page=12";
const raw = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const reference = (id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`;
const range = (xml, id = 0) => `<w:commentRangeStart w:id="${id}"/>${xml}<w:commentRangeEnd w:id="${id}"/>${reference(id)}`;
const para = (xml, id) => `<w:p${id ? ` w14:paraId="${id}"` : ""}>${xml}</w:p>`;
const comments = [
  { author: "Writer", content: raw(rootText) },
  { author: "Reviewer", content: raw("#判断 ") + `<w:hyperlink r:id="rId1">${raw("表3")}</w:hyperlink>` },
];

async function docx(paragraphs, { threads = [], resolved = false } = {}) {
  const zip = new JSZip();
  zip.file("word/document.xml", `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphs.join("")}</w:body></w:document>`);
  if (threads.length) {
    const ids = threads.map((_, i) => String(90000000 + i));
    zip.file("word/comments.xml", `<w:comments xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}">${threads.map((c, i) =>
      `<w:comment w:id="${i}" w:author="${c.author}"><w:p w14:paraId="${ids[i]}">${c.content}</w:p></w:comment>`).join("")}</w:comments>`);
    zip.file("word/commentsExtended.xml", `<w15:commentsEx xmlns:w15="${W15}">${ids.map((id, i) =>
      `<w15:commentEx w15:paraId="${id}" w15:done="${i === 0 && resolved ? 1 : 0}"${i ? ` w15:paraIdParent="${ids[0]}"` : ""}/>`).join("")}</w15:commentsEx>`);
    zip.file("word/commentsIds.xml", `<w16cid:commentsIds xmlns:w16cid="${IDS}">${ids.map((id, i) =>
      `<w16cid:commentId w16cid:paraId="${id}" w16cid:durableId="${String(10000000 + i)}"/>`).join("")}</w16cid:commentsIds>`);
    zip.file("word/_rels/comments.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="${R}/hyperlink" TargetMode="External" Target="https://example.org/another?report=a&amp;part=2#table3"/>
    </Relationships>`);
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

try {
  const bundle = path.join(tmp, "implementation.cjs");
  await esbuild.build({
    stdin: { resolveDir: here, contents: `
      export { readSelectedSavedVersions, readSavedVersionArchive } from "./src/savedLineage";
      export { LineageComments } from "./src/lineageComments";
      export { compareSavedVersions, lineageRangeText } from "../src/lib/lineage";
      export { commentReferences } from "../src/lib/commentReferences";
      export { extractComments, diffComments } from "../src/lib/docxComments";
      export { WatchArchive } from "../src/node/docWatcher";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile: bundle, logLevel: "silent",
  });
  const { readSelectedSavedVersions, readSavedVersionArchive, LineageComments, compareSavedVersions, lineageRangeText,
    commentReferences, extractComments, diffComments, WatchArchive } = require(bundle);
  const source = path.join(tmp, "report.docx");
  const archive = WatchArchive.open(path.join(tmp, "watch"), source);
  const quote = "欠席率が20％低下した";
  const before = "制度によって";
  const after = "制度の導入後に";
  const fixed = para(raw("前提と調査の手順を以下に示す。"), "00112233");
  const unrelated = (text) => para(raw(text), "55667788");
  const passage = (prefix, anchored) => para(
    raw(`📝${quote}の記述を再確認した。${prefix}`) + (anchored ? range(raw(quote)) : raw(quote)) + raw("。"),
    "22334455",
  );
  const timeline = [
    [[fixed, passage(before), unrelated("本文以外の注記その一。")], {}],
    [[fixed, passage(before, true), unrelated("本文以外の注記その一。")], { threads: comments }],
    [[passage(after, true), fixed, unrelated("本文以外の注記その一。")], { threads: comments, resolved: true }],
    [[passage(after), fixed, unrelated("本文以外の注記その二。")], {}],
    [[passage(after), fixed, unrelated("本文以外の注記その三。")], {}],
  ];
  const selected = [];
  for (const [index, [paragraphs, options]] of timeline.entries()) {
    const bytes = await docx(paragraphs, options);
    const file = path.join(tmp, `version-${index + 1}.docx`);
    const at = new Date(Date.UTC(2026, 0, 1 + index));
    await fs.writeFile(file, bytes);
    selected.push(file);
    const entry = await archive.addVersion(bytes, at, at);
    archive.appendEvents(diffComments(archive.commentState(), await extractComments(bytes), at, entry.sha256));
  }
  const manifest = path.join(archive.dir, "snapshots.json");
  const events = path.join(archive.dir, "comments.jsonl");
  const unchanged = await Promise.all([fs.readFile(manifest), fs.readFile(events)]);
  const archived = await readSavedVersionArchive(archive);
  const manual = await readSelectedSavedVersions(selected);
  assert.deepEqual(archived.map((version) => version.threads.length), [0, 1, 1, 0, 0]);
  assert.deepEqual(manual.map((version) => version.threads.length), [0, 1, 1, 0, 0]);
  const result = compareSavedVersions(archived, { japaneseWordHints: true });
  const index = new LineageComments(archived, result);
  const commentOnly = result.comparisons[0].paragraph.find((link) => link.to[0]?.paragraph === 1);
  assert.equal(commentOnly.kind, "retained");
  assert.equal(index.forRelationship(commentOnly, "paragraph")[0].version, 1);
  const unchangedInLaterSaves = result.comparisons[3].paragraph.find((link) => link.to[0]?.paragraph === 0);
  const found = index.forRelationship(unchangedInLaterSaves, "paragraph");
  assert.deepEqual(found.map((entry) => entry.version), [2, 1]);
  assert(found.every((entry) => index.absentInLatest(entry)));
  const [observed] = found;
  assert.equal(observed.thread.detail.version.n, 3);
  assert.deepEqual(observed.thread.detail.observations.map((e) => [e.type, e.version]), [["resolved", 3], ["removed", 4]]);
  assert.equal(observed.thread.detail.comments[1].author, "Reviewer");
  assert.deepEqual(commentReferences(observed.thread.detail.comments[0].text).tags, ["出典確認"]);
  assert.deepEqual(commentReferences(observed.thread.detail.comments[0].text).links.map((l) => l.url), ["https://example.org/report.pdf#page=12"]);
  assert.deepEqual(commentReferences(observed.thread.detail.comments[1].text, observed.thread.detail.comments[1].links), {
    tags: ["判断"], links: [{ label: "表3", url: "https://example.org/another?report=a&part=2#table3" }],
  });
  const [original] = observed.thread.ranges;
  const originalParagraph = archived[2].paragraphs[0].text;
  assert(original.start > originalParagraph.indexOf(quote), "use the Word marker, not the first matching string");
  assert.equal(lineageRangeText(archived, { ...original, version: observed.version }), quote);
  assert.equal(index.forRelationship(result.comparisons[3].paragraph.find((link) => link.to[0]?.paragraph === 1), "paragraph").length, 0);
  console.log("ok archive and manual versions, comment-only save, moved paragraph, removed thread, replies, tags, links, original range");

  assert.deepEqual(await fs.readFile(manifest), unchanged[0]);
  assert.deepEqual(await fs.readFile(events), unchanged[1]);
  await fs.unlink(path.join(archive.dir, archive.versions()[2].file));
  const backup = (await readSavedVersionArchive(archive))[2];
  assert(backup.unavailable);
  assert(backup.threads[0].detail.versionUnavailable);
  assert.equal(backup.threads[0].detail.recordedAnchor.quote, quote);
  assert.equal(backup.threads[0].ranges.length, 0, "recorded quote does not establish a text offset without the saved file");

  const lineA = "A complete paragraph about the independent evaluation.";
  const lineB = "An unrelated paragraph about the archival evidence.";
  const fixtures = [
    await docx([para(raw(lineA + " ") + range(raw(lineB)))], { threads: comments }),
    await docx([para(raw(lineA)), para(raw(lineB))]),
    await docx([para(raw(lineA)), para(raw(lineB)), para(raw("The review is now complete."))]),
  ];
  const splitPaths = await Promise.all(fixtures.map(async (bytes, i) => {
    const file = path.join(tmp, `split-${i}.docx`);
    await fs.writeFile(file, bytes);
    return file;
  }));
  const split = await readSelectedSavedVersions(splitPaths);
  const splitResult = compareSavedVersions(split);
  const splitIndex = new LineageComments(split, splitResult);
  assert(splitResult.comparisons[0].paragraph.some((link) => link.kind === "split"));
  const [first, second] = splitResult.comparisons[1].paragraph;
  assert.equal(splitIndex.forRelationship(first, "paragraph").length, 0);
  assert.equal(splitIndex.forRelationship(second, "paragraph").length, 1);

  const duplicatePaths = [];
  for (const [i, xml] of [
    await docx([para(range(raw(lineA))), para(raw(lineA))], { threads: comments }),
    await docx([para(raw(lineA)), para(raw(lineA))]),
  ].entries()) {
    const file = path.join(tmp, `uncertain-${i}.docx`);
    await fs.writeFile(file, xml);
    duplicatePaths.push(file);
  }
  const uncertain = await readSelectedSavedVersions(duplicatePaths);
  const uncertainResult = compareSavedVersions(uncertain);
  const uncertainIndex = new LineageComments(uncertain, uncertainResult);
  assert.equal(uncertain[0].threads[0].ranges[0].paragraph, 0);
  assert(uncertainResult.comparisons[0].paragraph.every((link) =>
    link.kind === "ambiguous" && uncertainIndex.forRelationship(link, "paragraph").length === 0));
  console.log("ok local records fallback, split targets and ambiguous counterparts never inherit neighboring comments");
} finally {
  await fs.rm(tmp, { recursive: true, force: true });
}
