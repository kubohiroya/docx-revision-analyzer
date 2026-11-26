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

async function fixture(index, withUndatedRevision = false) {
  const zip = new JSZip();
  const prefix = run(index >= 2 ? "制度導入後、" : "制度によって");
  const initial = withUndatedRevision ? '<w:ins w:id="91">' + prefix + '</w:ins>' : prefix;
  const focus = para("00000001", initial + (index === 1 ? anchor(quote) : run(quote)) + run("。"));
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
      export { compareSavedVersionPath, inferSavedVersionTimeline } from "./src/versionReviewPath";
      export { addSavedVersionComment, readSelectedSavedVersions, readSavedVersionArchive, savedVersionFile } from "./src/savedLineage";
      export { compareSavedVersions } from "../src/lib/lineage";
      export { extractComments, diffComments } from "../src/lib/docxComments";
      export { WatchArchive } from "../src/node/docWatcher";
      export { replaceWithBackup } from "../src/node/historyFile";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile: bundle,
  });
  const { SavedReportAnalysis, savedReportTitle, savedReportHtml, compareSavedVersionPath, inferSavedVersionTimeline,
    readSelectedSavedVersions, readSavedVersionArchive, savedVersionFile,
    addSavedVersionComment, compareSavedVersions, extractComments, diffComments, WatchArchive, replaceWithBackup } = createRequire(import.meta.url)(bundle);
  const meeting = "説明会は五号館で開催します。担当部署は回答の根拠を当日までにそろえます。";
  const originalBody = [
    meeting,
    "今月の説明会は事務局が司会を担当し、議事録は終了後に配布します。",
    "当日は参加者名簿を確認し、案内担当者と会場責任者の間で共有します。",
  ];
  const branchA = ["Ａ案では図書室を主会場として準備し、展示物を先に運びます。",
    "Ａ案の案内係は閲覧室で名簿を受け取り、全員に利用方法を説明します。"];
  const branchB = ["Ｂ案では配信設備を先に見直し、字幕担当者と操作手順を共有します。",
    "Ｂ案の収録班は通信状況を確認し、変更した機材を会議室に持ち込みます。"];
  function branchVersion(label, hour, body, key, notes = []) {
    return {
      label, source: "selected", sourceModified: new Date(Date.UTC(2026, 0, 1, hour)).toISOString(), trackedChanges: false,
      paragraphs: body.map((text, index) => ({ index, kind: "body", text, paraId: String(index + 1) })),
      threads: key ? [{ key, ranges: [{ paragraph: 0, start: 0, end: 14 }], detail: {
        versionUnavailable: false, observations: [],
        comments: [{ key, id: "0", done: false, author: key, text: key },
          ...notes.map((text, index) => ({ key: key + index, id: String(index + 1), done: false, text }))],
      } }] : [],
    };
  }
  const changedMeeting = meeting.replace("五号館", "六号館");
  const branchVersions = [
    branchVersion("root.docx", 0, originalBody, "ROOT_HISTORICAL", ["旧版の返信 #参照 https://example.org/root-source"]),
    branchVersion("A1.docx", 1, [...originalBody, ...branchA], "A_PERSISTS"),
    branchVersion("B1.docx", 2, [...originalBody, ...branchB], "SIBLING_PRIVATE", ["B_REPLY #B_BRANCH https://example.org/b-source"]),
    branchVersion("A2.docx", 3, [changedMeeting, ...originalBody.slice(1), ...branchA], "A_PERSISTS"),
    branchVersion("B2.docx", 4, [meeting.replace("五号館", "七号館"), ...originalBody.slice(1), ...branchB]),
    branchVersion("unrelated.docx", 5, ["配車担当者は各車両の燃料を確認し、整備先を一覧に記録します。"]),
  ];
  const { paths: allPaths } = inferSavedVersionTimeline(branchVersions, [0, 1, 2, 3, 4, 5]);
  assert.deepEqual(allPaths.get("0:1:3"), [0, 1, 3]);
  assert.deepEqual(allPaths.get("2:4"), [2, 4]);
  assert(!inferSavedVersionTimeline(branchVersions, [0, 2, 3, 4, 5]).paths.has("0:1:3"));

  const exportBranch = (analysis, pair, labels, extra = {}) => analysis.project({
    versions: [...pair], initial: pair, onlyPairs: [pair], range: 2,
    excluded: new Set(), includeText: true, includeComments: true, includeDeleted: true,
    includeLinks: true, includeUnlocated: false, includeRelatedSources: true, ...extra,
  }, "ja", labels);
  const labels = branchVersions.map(version => version.label);
  for (const level of ["paragraph", "sentence"]) {
    const chronological = new SavedReportAnalysis(branchVersions, compareSavedVersions(branchVersions), level);
    const unscoped = exportBranch(chronological, [1, 3], labels);
    assert(unscoped.threads.some(thread => thread.messages.some(message => message.text === "SIBLING_PRIVATE")));
    assert(unscoped.threads.some(thread => thread.absent && thread.messages[0].text === "A_PERSISTS"));

    const analysis = compareSavedVersionPath(branchVersions, [0, 1, 3], level);
    const scopeLabels = [labels[0], labels[1], labels[3]];
    const review = exportBranch(analysis, [1, 2], scopeLabels, {
      versions: [0, 1, 2], initial: [0, 2], onlyPairs: [[0, 1], [1, 2]], review: true,
    });
    assert.deepEqual(review.drafts.map(draft => draft.filename), scopeLabels);
    assert(review.pairs[1].changes.length);
    assert(review.threads.every(thread => !thread.messages.some(message => message.text === "SIBLING_PRIVATE")));

    const output = exportBranch(analysis, [1, 2], scopeLabels);
    assert.deepEqual(output.drafts.map(draft => draft.filename), [labels[1], labels[3]]);
    const messages = keyword => output.threads.filter(thread => thread.messages.some(message => message.text === keyword));
    assert(messages("A_PERSISTS").length);
    assert(messages("A_PERSISTS").every(thread => !thread.absent));
    const deleted = messages("ROOT_HISTORICAL").find(thread => thread.original?.length);
    assert(deleted?.absent);
    assert.equal(deleted.sourceLabel, "root.docx");
    assert.deepEqual(deleted.messages[1].tags, ["参照"]);
    assert.equal(deleted.messages[1].links[0].url, "https://example.org/root-source");
    const html = savedReportHtml(output);
    assert(html.includes("ROOT_HISTORICAL"));
    for (const term of ["SIBLING_PRIVATE", "B_REPLY", "B_BRANCH", "b-source", "B1.docx", "B2.docx", branchB[0]]) {
      assert(!html.includes(term), term);
    }
    const other = exportBranch(compareSavedVersionPath(branchVersions, [2, 4], level), [0, 1], [labels[2], labels[4]]);
    assert(other.threads.some(thread => thread.absent && thread.messages[0].text === "SIBLING_PRIVATE"));
    assert(!savedReportHtml(other).includes("A_PERSISTS"));
  }

  const archivedBranches = branchVersions.slice(0, 4).map((version, index) => ({
    ...version, source: "archive", number: index + 1, gapBefore: index === 2,
  }));
  const archivedPath = compareSavedVersionPath(archivedBranches, [0, 1, 3], "paragraph");
  assert(exportBranch(archivedPath, [1, 2], [labels[0], labels[1], labels[3]]).pairs[0].gap);
  console.log("ok inferred paths isolate comments and past sources in paragraph and sentence review and HTML");

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
  const undatedFile = path.join(temp, "undated.docx");
  await fs.writeFile(undatedFile, await fixture(0, true));
  const [undatedVersion] = await readSelectedSavedVersions([undatedFile]);
  assert.equal(undatedVersion.revisionCount, 1);
  const undatedReport = new SavedReportAnalysis([undatedVersion], compareSavedVersions([undatedVersion]), "paragraph");
  assert.equal(undatedReport.drafts("ja", ["undated"])[0].characterDelta, Array.from("制度によって").length);
  assert.equal(typeof undatedReport.drafts("en", ["undated"])[0].wordDelta, "number");
  const shortenedFile = path.join(temp, "shortened.docx");
  const shortened = new JSZip();
  shortened.file("word/document.xml", `<w:document xmlns:w="${W}"><w:body>${para("00000001",
    '<w:del w:id="1" w:date="2026-01-02T00:00:00Z"><w:r><w:delText xml:space="preserve">A long </w:delText></w:r></w:del>' +
    run("draft"))}</w:body></w:document>`);
  await fs.writeFile(shortenedFile, await shortened.generateAsync({ type: "nodebuffer" }));
  const [shortenedVersion] = await readSelectedSavedVersions([shortenedFile]);
  const shortenedReport = new SavedReportAnalysis([shortenedVersion], compareSavedVersions([shortenedVersion]), "paragraph");
  const shortenedDraft = shortenedReport.drafts("en", ["shortened"])[0];
  assert.deepEqual([shortenedDraft.characters, shortenedDraft.words, shortenedDraft.characterDelta, shortenedDraft.wordDelta],
    [5, 1, -7, -2]);
  assert.equal(manual[0].revisionCount, 0);
  assert.equal(manual[1].revisionCount, 0);
  assert.equal(archived[0].revisionCount, 0);
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
    assert.deepEqual(report.drafts.map((draft) => draft.revisionCount), [0, 0, 0]);
    assert(report.drafts.every((draft) => draft.characterDelta === undefined && draft.wordDelta === undefined));
    assert.deepEqual(report.pairs.map((pair) => [pair.before, pair.after]), [[0, 1], [0, 2], [1, 2]]);
    const commentOnly = report.pairs[0].changes.find((row) => row.kind === "comment");
    assert.deepEqual(commentOnly?.characterChanges, { added: 0, deleted: 0 });
    const rewritten = (report) => report.pairs[1].changes.find((row) => row.kind === "rewritten" && row.before[0]?.paragraph === 2);
    assert.deepEqual(rewritten(report)?.characterChanges, { added: 4, deleted: 4 });
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
    assert.deepEqual(rewritten(withoutLinks)?.characterChanges, rewritten(report)?.characterChanges);
    assert(!JSON.stringify(withoutLinks).includes("第一ページの周辺情報。"));
    assert.equal(analysis.project({ ...scope, includeDeleted: false }, "ja", labels).threads.length, 0);
    assert.equal(analysis.project({ ...scope, includeComments: false }, "ja", labels).threads.length, 0);
    const withoutText = analysis.project({ ...scope, includeText: false }, "ja", labels);
    assert.equal(withoutText.clips.length, 0);
    assert(rewritten(withoutText) && !("characterChanges" in rewritten(withoutText)));
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

  const writeFolder = path.join(temp, "comment-writes");
  await fs.mkdir(writeFolder);
  const targetFile = path.join(writeFolder, "review.docx");
  const targetZip = await JSZip.loadAsync(protectedBytes[0]);
  targetZip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>');
  const targetBytes = await targetZip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(targetFile, targetBytes);
  const [targetVersion] = await readSelectedSavedVersions([targetFile], true);
  const paragraphText = "制度によって" + quote + "。";
  const request = { author: "校閲者", text: "#根拠 報告書を参照\nhttps://example.org/table", range: {
    paragraphId: "00000001", paragraphText, start: 0, end: paragraphText.length,
  } };
  assert.deepEqual(await addSavedVersionComment(archived[0].openId, request), { ok: false, reason: "unavailable" });
  assert.deepEqual(await addSavedVersionComment(targetFile, request), { ok: false, reason: "unavailable" });
  const lock = path.join(writeFolder, "~$view.docx");
  await fs.writeFile(lock, "Word lock");
  assert.deepEqual(await addSavedVersionComment(targetVersion.openId, request), { ok: false, reason: "locked" });
  await fs.rm(lock);
  await fs.chmod(targetFile, 0o400);
  assert.deepEqual(await addSavedVersionComment(targetVersion.openId, request), { ok: false, reason: "readOnly" });
  await fs.chmod(targetFile, 0o600);
  assert.deepEqual(await addSavedVersionComment(targetVersion.openId, { ...request,
    range: { ...request.range, paragraphText: "Different passage" } }), { ok: false, reason: "unsupported" });
  assert.deepEqual(await fs.readFile(targetFile), targetBytes);
  assert.deepEqual(await fs.readdir(writeFolder), ["review.docx"]);

  const [saved, duplicate] = await Promise.all([addSavedVersionComment(targetVersion.openId, request),
    addSavedVersionComment(targetVersion.openId, request)]);
  assert(saved.ok);
  assert.deepEqual(duplicate, { ok: false, reason: "changed" });
  assert.deepEqual(await fs.readFile(saved.backupPath), targetBytes);
  const [note] = await extractComments(await fs.readFile(targetFile));
  assert.equal(note.author, request.author);
  assert.equal(note.text, request.text);
  assert.equal(note.anchor.quote, paragraphText);
  assert(saved.version.openId && saved.version.openId !== targetVersion.openId);
  const [reloaded] = await readSelectedSavedVersions([targetFile], true);
  assert.equal(reloaded.threads.length, 1);
  const savedBytes = await fs.readFile(targetFile);
  await assert.rejects(replaceWithBackup(targetFile, targetBytes, async () => { throw new Error("simulated conflict before rename"); }), /simulated conflict/u);
  assert.deepEqual(await fs.readFile(targetFile), savedBytes);
  assert(!(await fs.readdir(writeFolder)).some((name) => name.startsWith(".dra-replace-")));
  await fs.appendFile(targetFile, "An external modification");
  const changedBytes = await fs.readFile(targetFile);
  assert.deepEqual(await addSavedVersionComment(saved.version.openId, request), { ok: false, reason: "changed" });
  assert.deepEqual(await fs.readFile(targetFile), changedBytes);
  // Exercise a second file that already has a thread, a reply, and a tracked revision.
  // This is an OOXML regression test, not a substitute for opening the result in Word.
  const prior = await JSZip.loadAsync(await fixture(1, true));
  prior.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
  prior.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="${R}/comments" Target="comments.xml"/></Relationships>`);
  const priorBytes = await prior.generateAsync({ type: "nodebuffer" });
  const threadedFile = path.join(writeFolder, "threaded.docx");
  await fs.writeFile(threadedFile, priorBytes);
  const [threaded] = await readSelectedSavedVersions([threadedFile], true);
  const oldNotes = await extractComments(priorBytes);
  assert(oldNotes.some(note => note.parentKey && oldNotes.some(parent => parent.key === note.parentKey && parent.done)));
  const richText = '2行目を確認 & <確認済み> / verify “terms”\n#根拠 https://example.org/source';
  const untouchedParagraph = pageContext + "第一ページの周辺情報。";
  const targetQuote = "授業の記録";
  const posted = await addSavedVersionComment(threaded.openId, {
    author: "Review / 校閲", text: richText,
    range: { paragraphId: "00000002", paragraphText: untouchedParagraph, start: 0, end: targetQuote.length },
  });
  assert(posted.ok, JSON.stringify(posted));
  assert.deepEqual(await fs.readFile(posted.backupPath), priorBytes);
  const written = await fs.readFile(threadedFile);
  const newNotes = await extractComments(written);
  assert.deepEqual(newNotes.slice(0, oldNotes.length), oldNotes, "existing comments, reply links and resolution stay intact");
  assert.equal(newNotes.length, oldNotes.length + 1);
  assert.equal(newNotes.at(-1).text, richText);
  assert.equal(newNotes.at(-1).author, "Review / 校閲");
  assert.equal(newNotes.at(-1).parentKey, undefined);
  assert.equal(newNotes.at(-1).anchor.quote, targetQuote);
  assert.deepEqual(posted.version.paragraphs, threaded.paragraphs);
  assert.equal(posted.version.revisionCount, threaded.revisionCount);
  console.log("ok comment writes: opaque local capabilities, lock, protection, conflict, backup, serialized submission and reload");
} finally {
  await fs.rm(temp, { recursive: true, force: true });
}
