import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const temp = await fs.mkdtemp(path.join(os.tmpdir(), "dra-saved-report-"));
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const direct = "https://example.org/report-direct?a=b#page=12";
const target = "https://example.org/report-labelled?part=a&table=3#results";
const escaped = (text) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const run = (text) => `<w:r><w:t xml:space="preserve">${escaped(text)}</w:t></w:r>`;
const para = (id, content) => `<w:p w14:paraId="${id}">${content}</w:p>`;
const quote = "欠席率が20％低下した";
const dangerous = '</script><img src="https://example.org/report-pixel">';
const pageContext = "授業の記録を調べ、対象校について学期ごとの出欠を照合した。".repeat(17);
const anchor = (text) => `<w:commentRangeStart w:id="0"/>${run(text)}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`;

async function fixture(index) {
  const zip = new JSZip();
  const focus = para("00000001", run(index >= 2 ? "制度導入後、" : "制度によって") + (index === 1 ? anchor(quote) : run(quote)) + run("。"));
  const context = para("00000002", run(pageContext + "第一ページの周辺情報。"));
  const secret = para("00000003", `<w:r><w:lastRenderedPageBreak/></w:r>${run("第二ページの収録対象外。")}`);
  const unsafe = para("00000004", run(index >= 2 ? dangerous : "安全な原文から始める。"));
  const excluded = para("00000005", run(index === 3 ? "選ばない稿の秘匿本文。".repeat(5) : "継続する注記。"));
  const body = index >= 2 ? [focus, context, unsafe, secret, excluded] : [context, focus, unsafe, secret, excluded];
  zip.file("word/document.xml", `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body.join("")}</w:body></w:document>`);
  if (index !== 2) zip.file("docProps/app.xml", "<Properties><Pages>2</Pages></Properties>");
  if (index === 1) {
    zip.file("word/comments.xml", `<w:comments xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}">
      <w:comment w:id="0" w:author="Writer"><w:p w14:paraId="90000000">${run("#出典確認 検証結果 " + direct)}</w:p></w:comment>
      <w:comment w:id="1" w:author="Reviewer"><w:p w14:paraId="90000001">${run("#判断 ")}
        <w:hyperlink r:id="rId1">${run("表3")}</w:hyperlink></w:p></w:comment>
      <w:comment w:id="2" w:author="Archived editor"><w:p w14:paraId="90000002">${run("位置未確定の注記")}</w:p></w:comment>
    </w:comments>`);
    zip.file("word/commentsExtended.xml", `<w15:commentsEx xmlns:w15="${W15}">
      <w15:commentEx w15:paraId="90000000" w15:done="1"/><w15:commentEx w15:paraId="90000001" w15:paraIdParent="90000000"/>
    </w15:commentsEx>`);
    zip.file("word/_rels/comments.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="${R}/hyperlink" TargetMode="External" Target="${escaped(target)}"/>
    </Relationships>`);
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

try {
  const viewer = await esbuild.build({
    entryPoints: [path.join(here, "src/renderer/savedReportEntry.ts")], bundle: true, platform: "browser",
    format: "iife", target: "es2022", minify: true, write: false,
  });
  const css = await fs.readFile(path.join(here, "src/renderer/savedReport.css"), "utf8");
  const bundle = path.join(temp, "implementation.cjs");
  await esbuild.build({
    define: {
      __DRA_SAVED_REPORT_SCRIPT__: JSON.stringify(viewer.outputFiles[0].text),
      __DRA_SAVED_REPORT_CSS__: JSON.stringify(css),
    },
    stdin: { resolveDir: here, contents: `
      export { SavedReportAnalysis, savedReportTitle } from "./src/savedReport";
      export { savedReportHtml } from "./src/renderer/savedReportHtml";
      export { readSelectedSavedVersions, readSavedVersionArchive, savedVersionFile } from "./src/savedLineage";
      export { compareSavedVersions } from "../src/lib/lineage";
      export { extractComments, diffComments } from "../src/lib/docxComments";
      export { WatchArchive } from "../src/node/docWatcher";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile: bundle,
  });
  const { SavedReportAnalysis, savedReportTitle, savedReportHtml, readSelectedSavedVersions, readSavedVersionArchive, savedVersionFile,
    compareSavedVersions, extractComments, diffComments, WatchArchive } = createRequire(import.meta.url)(bundle);
  const source = path.join(temp, "original.docx");
  const archive = WatchArchive.open(path.join(temp, "watch"), source);
  const fixtures = [];
  for (const index of [0, 1, 2, 3]) {
    const bytes = await fixture(index);
    const file = path.join(temp, index === 3 ? "excluded-draft.docx" : `draft-${index + 1}.docx`);
    await fs.writeFile(file, bytes);
    fixtures.push(file);
    const date = new Date(Date.UTC(2026, 0, index + 1));
    const entry = await archive.addVersion(bytes, date, date);
    archive.appendEvents(diffComments(archive.commentState(), await extractComments(bytes), date, entry.sha256));
  }
  const protectedPaths = [...fixtures, path.join(archive.dir, "snapshots.json"), path.join(archive.dir, "comments.jsonl")];
  const protectedBytes = await Promise.all(protectedPaths.map((file) => fs.readFile(file)));
  const manual = await readSelectedSavedVersions(fixtures, true);
  const archived = await readSavedVersionArchive(archive, true);
  assert.equal(await savedVersionFile(manual[0].openId), await fs.realpath(fixtures[0]));
  await assert.rejects(savedVersionFile(fixtures[0]), /no longer available/);
  const archiveCopy = await savedVersionFile(archived[0].openId);
  try {
    assert.notEqual(archiveCopy, path.join(archive.dir, archived[0].label));
    assert.deepEqual(await fs.readFile(archiveCopy), protectedBytes[0]);
  } finally {
    await fs.rm(path.dirname(archiveCopy), { recursive: true, force: true });
  }
  assert.equal(savedReportTitle(["出席支援制度の評価報告-初稿.docx", "出席支援制度の評価報告-中間稿.DOCX"]), "出席支援制度の評価報告");
  assert.equal(savedReportTitle(["報告-s01.docx", "報告-s12.docx"]), "報告");
  assert.equal(savedReportTitle(["2026-11-24_report.docx", "2026-11-25_report.docx"]), "report");
  assert.equal(savedReportTitle(["初稿.docx", "中間稿.docx", "最終稿.docx"]), "初稿 / 中間稿 / 最終稿");
  const info = await fs.stat(fixtures[0]);
  assert.equal(manual[0].fileSize, info.size);
  assert.equal(manual[0].fileCreatedAt, info.birthtimeMs > 0 ? info.birthtime.toISOString() : undefined);
  assert.equal(archived[0].fileSize, info.size);
  assert.equal(archived[0].fileCreatedAt, undefined);

  for (const versions of [manual, archived]) {
    assert.deepEqual(versions[1].pageMap, { breaks: [{ paragraph: 3, offset: 0 }], estimated: false });
    assert.equal(versions[2].pageMap.estimated, true);
    const lineage = compareSavedVersions(versions, { japaneseWordHints: true });
    const analysis = new SavedReportAnalysis(versions, lineage, "paragraph");
    const scope = {
      versions: [0, 1, 2], initial: [0, 1], excluded: new Set(), includeText: true, range: 4,
      includeComments: true, includeDeleted: true, includeLinks: true, includeUnlocated: false,
    };
    const labels = ["初稿", "中間稿", "最終稿", "収録対象外の稿"];
    const report = analysis.project(scope, "ja", labels);
    assert.equal(report.title, versions === manual ? "draft" : "original");
    assert.deepEqual(report.pairs.map((pair) => [pair.before, pair.after]), [[0, 1], [0, 2], [1, 2]]);
    assert(report.pairs[0].changes.some((row) => row.kind === "comment"));
    assert(report.pairs[1].changes.some((row) => row.kind === "rewritten" && row.evidence.includes("paragraphId")));
    const thread = report.threads.find((item) => item.draft === 1);
    assert(thread?.absent && thread.resolved);
    assert.deepEqual(thread.messages.map((message) => message.tags), [["出典確認"], ["判断"]]);
    assert.deepEqual(thread.messages[0].links, [{ url: direct, label: direct }]);
    assert.deepEqual(thread.messages[1].links, [{ url: target, label: "表3" }]);
    assert.equal(thread.original.length, 1);
    const { clip, marks: [[start, end]] } = thread.original[0].excerpt;
    assert.equal(report.clips[clip].text.slice(start, end), quote);
    assert(report.clips[clip].text.includes("第一ページの周辺情報。"));
    assert(!report.clips[clip].text.includes("第二ページの収録対象外。"));
    assert.deepEqual(report.clips[clip].page, { number: 1, estimated: false });
    for (const value of [temp, "収録対象外の稿", "選ばない稿の秘匿本文", "第二ページの収録対象外"]) {
      assert(!JSON.stringify(report).includes(value), value);
    }
    const html = savedReportHtml(report);
    assert(html.includes("<title>" + report.title + "</title>"));
    assert(savedReportHtml({ ...report, title: "A </title> & B" }).includes("<title>A &lt;/title&gt; &amp; B</title>"));
    assert(!html.includes(dangerous));
    assert(!html.includes('<img src="https://example.org/report-pixel"'));
    assert(html.includes("default-src 'none'"));
    assert(html.includes('type="application/json"'));
    const payload = html.match(/id="report-data" type="application\/json">([^]*?)<\/script>/u)[1];
    assert.equal(JSON.parse(payload).clips.some((item) => item.text.includes(dangerous)), true);

    const withoutLinks = analysis.project({ ...scope, includeLinks: false, range: 0 }, "ja", labels);
    assert(withoutLinks.threads.some((item) => item.messages[1].author === "Reviewer"));
    assert(!JSON.stringify(withoutLinks).includes(direct));
    assert(!JSON.stringify(withoutLinks).includes(target));
    assert(withoutLinks.clips.every((item) => item.text.length <= 20));
    assert(!JSON.stringify(withoutLinks).includes("第一ページの周辺情報。"));
    assert.equal(analysis.project({ ...scope, includeDeleted: false }, "ja", labels).threads.length, 0);
    assert.equal(analysis.project({ ...scope, includeComments: false }, "ja", labels).threads.length, 0);
    assert.equal(analysis.project({ ...scope, includeText: false }, "ja", labels).clips.length, 0);
    assert(!JSON.stringify(report).includes("位置未確定の注記"));
    const includeUnlocated = analysis.project({ ...scope, includeUnlocated: true }, "ja", labels);
    const [unlocated] = includeUnlocated.drafts[1].unlocated;
    assert.deepEqual(includeUnlocated.threads[unlocated].messages.map((m) => m.text), ["位置未確定の注記"]);
    assert.equal(includeUnlocated.threads[unlocated].original, undefined);
    const excluded = new Set(analysis.changes(0, 2).map((change) => change.id));
    assert.equal(analysis.project({ ...scope, excluded }, "ja", labels).pairs[1].changes.length, 0);
    const excludeMarkup = new Set(analysis.changes(0, 2).filter((change) => analysis.describe(change).includes("report-pixel")).map((change) => change.id));
    assert(excludeMarkup.size > 0);
    const small = analysis.project({ ...scope, versions: [0, 2], initial: [0, 2], range: 0, excluded: excludeMarkup }, "ja", labels);
    assert(!savedReportHtml(small).includes("report-pixel"));

    const related = analysis.changes(0, 2).find(({ comments }) => comments.some(({ version }) => version === 1));
    assert(related);
    const others = new Set(analysis.changes(0, 2).filter(({ id }) => id !== related.id).map(({ id }) => id));
    const selectedScope = { ...scope, versions: [0, 2], initial: [0, 2], range: 2,
      excluded: others, includeRelatedSources: true };
    const selectedReport = analysis.project(selectedScope, "ja", labels);
    const recordedElsewhere = selectedReport.threads.find((item) => item.sourceLabel === "中間稿");
    assert.deepEqual(selectedReport.drafts.map(({ label }) => label), ["初稿", "最終稿"]);
    assert.equal(selectedReport.pairs[0].changes.length, 1);
    assert(recordedElsewhere?.absent && recordedElsewhere.original?.length);
    assert.equal(recordedElsewhere.messages[1].author, "Reviewer");
    assert(selectedReport.drafts.every(({ filename, characters, words }) => filename && characters > 0 && words > 0));
    const selectedJson = JSON.stringify(selectedReport);
    for (const secret of [temp, versions[0].openId, "reviewId", "reviewVersion", "位置未確定の注記", "report-pixel"]) {
      assert(!selectedJson.includes(secret), secret);
    }
    const review = analysis.project({ ...selectedScope, review: true }, "ja", labels);
    assert.equal(review.pairs[0].changes[0].reviewId, related.id);
    assert.equal(review.threads.find((item) => item.sourceLabel === "中間稿").reviewVersion, 1);
  }
  assert.deepEqual(await Promise.all(protectedPaths.map((file) => fs.readFile(file))), protectedBytes);
  await fs.writeFile(fixtures[0], protectedBytes[1]);
  await assert.rejects(savedVersionFile(manual[0].openId), /has changed/);
  const unavailable = { label: "Unavailable", source: "archive", unavailable: true, threads: [] };
  const noPage = { ...manual[2], pageMap: undefined };
  const partial = [manual[0], unavailable, noPage];
  const result = new SavedReportAnalysis(partial, compareSavedVersions(partial), "paragraph").project({
    versions: [0, 1, 2], initial: [0, 2], excluded: new Set(), includeText: true, range: 4,
    includeComments: false, includeDeleted: false, includeLinks: false, includeUnlocated: false,
  }, "en", ["Before", "Missing", "After"]);
  assert(result.pairs[0].unavailable && result.pairs[2].unavailable);
  assert(result.pairs[1].gap && !result.pairs[1].unavailable);
  assert(result.clips.some((item) => !item.page && item.text.length <= 400));
  console.log("ok saved HTML: manual and archived drafts, all pairs, page bounds, deleted replies and origins, opt-in URLs, scope and escaping");
} finally {
  await fs.rm(temp, { recursive: true, force: true });
}
