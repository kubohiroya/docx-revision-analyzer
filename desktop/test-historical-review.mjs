import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, rm, readdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";
import JSZip from "jszip";

const here = dirname(fileURLToPath(import.meta.url));
const folder = await mkdtemp(join(tmpdir(), "dra-historical-review-"));
const node = (text) => `<w:r><w:t>${text}</w:t></w:r>`;
const ins = (text, id = 8) => `<w:ins w:id="${id}" w:author="編集者" w:date="2026-11-25T12:34:00Z">${node(text)}</w:ins>`;
const para = (content, id = "AB000010") => `<w:p w14:paraId="${id}">${content}</w:p>`;
async function docx(body, comment, parts = {}) {
  const zip = new JSZip();
  zip.file("word/document.xml",
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}</w:body></w:document>`);
  if (comment) zip.file("word/comments.xml",
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="4" w:author="Reader">${para(node(comment))}</w:comment></w:comments>`);
  for (const [name, content] of Object.entries(parts)) zip.file(name, content);
  return zip.generateAsync({ type: "nodebuffer" });
}

try {
  const compiled = join(folder, "tests.cjs");
  await esbuild.build({
    stdin: { resolveDir: here, contents: `
      export { readCloudSavedVersion, readSelectedSavedVersions, readSavedVersionRevisions } from "./src/savedLineage";
      export { groupedHistoryRows, historyEntries, historyMatch, reconcileHistoryVersions, searchKey } from "./src/historySearch";
      export { compareSavedVersions, lineageRangeText } from "../src/lib/lineage";
      export { projectHistoricalAnchor } from "../src/lib/historicalAnchors";
      export { listCloudVersionItems, getCloudVersionBytes, CloudVersionError } from "./src/cloudVersionGraph";
      export { CloudVersionImport } from "./src/cloudVersionImport";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile: compiled,
  });
  const {
    readCloudSavedVersion, readSelectedSavedVersions, readSavedVersionRevisions, groupedHistoryRows, historyEntries, historyMatch, reconcileHistoryVersions, searchKey,
    compareSavedVersions, lineageRangeText, projectHistoricalAnchor, listCloudVersionItems, getCloudVersionBytes, CloudVersionImport, CloudVersionError,
  } = createRequire(import.meta.url)(compiled);

  const left = "説明の前置き📝", right = "を検証しました。";
  const deletion = '<w:del w:id="12" w:author="Reader"><w:r><w:delText>旧案</w:delText></w:r></w:del>';
  const start = '<w:commentRangeStart w:id="4"/>';
  const end = '<w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r>';
  const before = await docx(para(node(left) + deletion + start + ins("新案") + end + node(right)) +
    para(node("First ") + ins("draft"), "AB000020") + para(node("Second ") + ins("draft"), "AB000030"), "中間版のコメント #資料");
  const after = await docx(para(node(left + "新案" + right)) +
    para(node("First draft"), "AB000020") + para(node("Second draft"), "AB000030"));
  const files = ["before.docx", "after.docx", "duplicate.docx"].map(name => join(folder, name));
  await Promise.all([writeFile(files[0], before), writeFile(files[1], after), writeFile(files[2], before)]);
  const versions = await readSelectedSavedVersions(files, true);
  assert.notEqual(versions[0].versionId, versions[2].versionId);
  assert.equal(versions[0].sha256, versions[2].sha256);
  const read = await readSavedVersionRevisions(versions[0].openId);
  assert.equal(read.status, "read");
  assert.deepEqual(await readSavedVersionRevisions(versions[0].openId, false), { status: "disabled", records: [] });
  assert.deepEqual(read.excluded, []);
  assert.equal(read.records[1].author, "編集者");
  assert.equal(read.records[1].date, "2026-11-25T12:34:00.000Z");
  assert.deepEqual(read.records.map(entry => [entry.kind, entry.text]), [
    ["del", "旧案"], ["ins", "新案"], ["ins", "draft"], ["ins", "draft"],
  ]);
  assert.equal(new Set(read.records.map(record => record.occurrence)).size, 4);
  assert.equal(read.records[0].date, undefined);
  assert.deepEqual(read.records[0].ranges, [{ paragraph: 0, start: left.length, end: left.length }]);
  assert.deepEqual((await readSavedVersionRevisions(versions[2].openId)).records, read.records);
  const history = historyEntries(versions[0], 0, read.records);
  assert.equal(history.find(entry => entry.kind === "comment").quote, "新案");
  assert.deepEqual(history.find(entry => entry.kind === "comment").tags, ["資料"]);
  assert.deepEqual(historyMatch("📝カ\u3099ABC", searchKey("ガa")), { start: 2, end: 5 });
  assert.equal(history.find(entry => entry.text === "旧案").version, 0);
  const lineage = compareSavedVersions(versions.slice(0, 2));
  assert.equal(projectHistoricalAnchor(lineage, 0, 1, read.records[0].ranges).kind, "boundary");
  assert.equal(projectHistoricalAnchor(lineage, 0, 1, read.records[1].ranges).kind, "range");
  const gap = compareSavedVersions([versions[0], { ...versions[1], gapBefore: true }]);
  assert.equal(projectHistoricalAnchor(gap, 0, 1, read.records[1].ranges).reason, "gap");
  const copy = compareSavedVersions([versions[0], { ...versions[1],
    paragraphs: [...versions[1].paragraphs, { ...versions[1].paragraphs[0], index: 10 }],
  }]);
  assert.equal(projectHistoricalAnchor(copy, 0, 1, read.records[1].ranges).kind, "ambiguous");
  const move = (side, group, text) =>
    `<w:${side}RangeStart w:id="81" w:name="${group}"/><w:${side} w:id="82">${node(text)}</w:${side}><w:${side}RangeEnd w:id="81"/>`;
  const rich = await docx(
    para(node("本文") + '<w:ins w:id="9">' + node("📝") + deletion +
      '<w:r><w:br/></w:r>' + node("ABC") + '</w:ins>') +
    para(move("moveFrom", "paired", "移動"), "AB000020") +
    para(move("moveTo", "paired", "移動") + move("moveTo", "unpaired", "未対応"), "AB000030"),
  );
  const richPath = join(folder, "rich.docx");
  await writeFile(richPath, rich);
  const richVersion = (await readSelectedSavedVersions([richPath], true))[0];
  const richIndex = await readSavedVersionRevisions(richVersion.openId);
  assert.deepEqual(richIndex.records.map(record => [record.kind, record.text]), [
    ["ins", "📝旧案\nABC"], ["del", "旧案"], ["moveFrom", "移動"], ["moveTo", "移動"], ["moveTo", "未対応"],
  ]);
  assert.deepEqual(richIndex.records[0].ranges, [{ paragraph: 0, start: 2, end: 4 }, { paragraph: 0, start: 4, end: 8 }]);
  assert.deepEqual(richIndex.records[1].ranges, [{ paragraph: 0, start: 4, end: 4 }]);
  assert.deepEqual(richIndex.records.slice(2).map(record => record.moveOccurrence), [3, 2, undefined]);
  assert.deepEqual(await readSavedVersionRevisions(richVersion.openId), richIndex);
  assert.deepEqual(await readFile(richPath), rich, "indexing never changes its source DOCX");

  const outside = await docx(
    '<w:tbl><w:tr><w:tc>' + para(ins("Table edit")) + '</w:tc></w:tr></w:tbl>' +
    para('<w:pPr><w:pPrChange w:id="77"><w:pPr/></w:pPrChange></w:pPr>' +
      '<w:r><w:pict><w:txbxContent>' + para(ins("Text box edit")) + '</w:txbxContent></w:pict></w:r>'),
    undefined, { "word/header1.xml": "<w:hdr/>", "word/footnotes.xml": "<w:footnotes/>" },
  );
  const readBytes = async (bytes, versionId) => {
    const version = await readCloudSavedVersion(bytes, { label: versionId, source: "cloud", versionId }, async () => bytes);
    return readSavedVersionRevisions(version.openId);
  };
  const outsideIndex = await readBytes(outside, "coverage");
  assert.deepEqual(outsideIndex, { status: "read", records: [], excluded: ["tables", "formatting", "textBoxes", "headersFooters", "notes"] });
  assert.deepEqual(await readBytes(outside, "same bytes in another version"), outsideIndex, "the cache preserves coverage");
  assert.deepEqual(await readBytes(await docx("<w:p></w:r>"), "bad XML"), { status: "failed", records: [] });
  await rm(richPath);
  assert.deepEqual(await readSavedVersionRevisions(richVersion.openId), { status: "unavailable", records: [] });

  const removedPath = join(folder, "removed.docx");
  const remaining = para(node("First draft"), "AB000020") + para(node("Second draft"), "AB000030");
  await writeFile(removedPath, await docx(remaining));
  const removed = (await readSelectedSavedVersions([removedPath], true))[0];
  const three = [versions[0], removed, versions[1]];
  const returned = compareSavedVersions(three);
  const originalRange = history.find(entry => entry.kind === "comment").ranges;
  const deleted = projectHistoricalAnchor(returned, 0, 1, originalRange);
  assert.deepEqual(deleted, { kind: "deleted", ranges: [] });
  const revived = projectHistoricalAnchor(returned, 0, 2, originalRange);
  assert.deepEqual(revived, { kind: "reappeared", ranges: [{ paragraph: 0, version: 2, start: 0, end: (left + "新案" + right).length }] });
  const anotherParagraph = projectHistoricalAnchor(returned, 1, 2, [{ paragraph: 0, start: 0, end: 5 }]);
  assert.deepEqual(anotherParagraph, { kind: "range", ranges: [{ paragraph: 1, version: 2, start: 0, end: 5 }] },
    "a past link with the same paragraph number is unrelated to a different version");
  const uncertain = compareSavedVersions([versions[0], removed, {
    ...versions[1], paragraphs: [...versions[1].paragraphs, { ...versions[1].paragraphs[0], index: 10 }],
  }]);
  assert.equal(projectHistoricalAnchor(uncertain, 0, 2, originalRange).kind, "ambiguous");
  const unavailable = compareSavedVersions([versions[0], { ...removed, unavailable: true }, versions[1]]);
  assert.equal(projectHistoricalAnchor(unavailable, 0, 2, originalRange).reason, "gap");
  const skipped = compareSavedVersions([versions[0], { ...removed, gapBefore: true }, versions[1]]);
  assert.equal(projectHistoricalAnchor(skipped, 0, 2, originalRange).reason, "gap");
  assert.equal(projectHistoricalAnchor(lineage, 0, 1, [{ paragraph: 0, start: 0, end: 0 }]).reason, "no-path");

  const spans = [];
  const first = "The evaluation used three independent tests.";
  const second = "We archived the raw observations for future readers.";
  const movedFrom = "制度によって欠席率が20％低下した。この結果は別の調査でも同じだった。";
  const movedTo = "制度の導入後、欠席率が20％低下した。この結果は別の調査でも同じだった。";
  const lead = "すべての調査票は別のフォルダに保管した。";
  for (const [i, frag] of [
    para(node(movedFrom), "AB000040") + para(node(lead), "AB000050") + para(node(first + " " + second), "AB000060"),
    para(node(lead), "AB000050") + para(node(movedTo), "AB000040") + para(node(first + " " + second), "AB000060"),
    para(node(lead), "AB000050") + para(node(movedTo), "AB000040") + para(node(first), "AB000070") + para(node(second), "AB000080"),
    para(node(lead), "AB000050") + para(node(movedTo), "AB000040") + para(node(first + " " + second), "AB000060"),
  ].entries()) {
    const file = join(folder, `span-${i}.docx`);
    await writeFile(file, await docx(frag));
    spans.push(file);
  }
  const changes = compareSavedVersions(await readSelectedSavedVersions(spans, true));
  const retainedText = "欠席率が20％低下した";
  const shifted = projectHistoricalAnchor(changes, 0, 3, [{
    paragraph: 0, start: movedFrom.indexOf(retainedText), end: movedFrom.indexOf(retainedText) + retainedText.length,
  }]);
  assert.equal(shifted.kind, "range");
  assert.deepEqual(shifted.ranges.map(range => [range.paragraph, lineageRangeText(changes.versions, range)]), [[1, retainedText]]);
  assert.equal(projectHistoricalAnchor(changes, 0, 1, [{ paragraph: 0, start: 2, end: 4 }]).kind, "related");
  const selected = "raw observations";
  for (const target of [2, 3]) {
    const start = first.length + 1 + second.indexOf(selected);
    const match = projectHistoricalAnchor(changes, 1, target, [{ paragraph: 2, start, end: start + selected.length }]);
    assert.equal(match.kind, "range");
    assert.deepEqual(match.ranges.map(range => [range.paragraph, lineageRangeText(changes.versions, range)]), [[target === 2 ? 3 : 2, selected]]);
  }
  const multiPath = join(folder, "multi.docx");
  await writeFile(multiPath, await docx(para(node("Lead ") + start + node("first"), "AB000010") +
    para(node("second") + end + node(" tail"), "AB000020"), "Span"));
  const multi = (await readSelectedSavedVersions([multiPath], true))[0];
  assert.equal(multi.threads[0].ranges.length, 2);
  const multiLineage = compareSavedVersions([multi, { ...multi, threads: [] }]);
  const multiMatch = projectHistoricalAnchor(multiLineage, 0, 1, multi.threads[0].ranges);
  assert.deepEqual(multiMatch.ranges.map(range => lineageRangeText(multiLineage.versions, range)), ["first", "second"]);
  const onlyFirst = compareSavedVersions([multi, { ...multi, paragraphs: multi.paragraphs.slice(0, 1), threads: [] }]);
  const surviving = projectHistoricalAnchor(onlyFirst, 0, 1, multi.threads[0].ranges);
  assert.equal(surviving.kind, "partial");
  assert.deepEqual(surviving.ranges.map(range => lineageRangeText(onlyFirst.versions, range)), ["first"]);

  const prefix = "Shared context before the edit. ";
  const suffix = " Shared context after the edit.";
  const long = (text) => ({ ...multi, paragraphs: [{ index: 0, paraId: "FEDCBA01", kind: "body", text }], threads: [] });
  const longMatch = compareSavedVersions([long(prefix + "a".repeat(600) + suffix), long(prefix + "b".repeat(600) + suffix)]);
  assert(longMatch.comparisons[0].paragraph.some(link => link.alignmentPartial));
  assert.equal(projectHistoricalAnchor(longMatch, 0, 1, [{ paragraph: 0, start: prefix.length + 30, end: prefix.length + 35 }]).kind, "related");
  assert.equal(projectHistoricalAnchor(longMatch, 0, 1, [{ paragraph: 0, start: prefix.length + 30, end: prefix.length + 30 }]).reason, "no-path");

  await writeFile(files[0], after);
  assert.equal((await readSavedVersionRevisions(versions[0].openId)).status, "changed");
  assert.equal((await readSavedVersionRevisions(versions[0].openId, false)).status, "changed");
  console.log("ok observed markup with identical visible text, duplicate IDs, exact source ranges, Unicode search and safe reconnection");

  function threadParts(text, { reply, resolved = false } = {}) {
    const contents = [text, ...(reply ? [reply] : [])];
    const ids = ["AA000004", "AA000005"];
    const xml = (tag, child, render) => contents.map((value, i) => render(`<${tag}:${child}`, value, i, ids[i])).join("");
    return {
      "word/comments.xml": '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
        xml("w", "comment", (tag, value, i, id) => `${tag} w:id="${i + 4}" w:author="${i ? "Reviewer" : "Reader"}">${para(value, id)}</w:comment>`) + "</w:comments>",
      "word/commentsExtended.xml": '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
        xml("w15", "commentEx", (tag, _value, i, id) => `${tag} w15:paraId="${id}"${i ? ' w15:paraIdParent="AA000004"' : ` w15:done="${Number(resolved)}"`}/>`) + "</w15:commentsEx>",
      "word/commentsIds.xml": '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">' +
        xml("w16cid", "commentId", (tag, _value, i, id) => `${tag} w16cid:paraId="${id}" w16cid:durableId="CC00000${i}"/>`) + "</w16cid:commentsIds>",
    };
  }
  const visible = para(node(left + "新案" + right));
  const commented = para(node(left) + start + node("新案") + end + node(right));
  const oldParts = threadParts(node("旧版の指摘[原文]") + '<w:r><w:br/></w:r>' + node("next 📝"), { reply: node("中間版だけの返信 #根拠") });
  const newParts = threadParts(node("編集後の指摘"), { resolved: true });
  const commentPaths = [];
  for (const [index, [body, parts]] of [[visible], [commented, oldParts], [commented, newParts], [visible], [commented]].entries()) {
    const filename = join(folder, `thread-${index}.docx`);
    await writeFile(filename, await docx(body, undefined, parts));
    commentPaths.push(filename);
  }
  const commentVersions = await readSelectedSavedVersions(commentPaths, true);
  assert(commentVersions.every(version => version.paragraphs[0].text === commentVersions[0].paragraphs[0].text));
  const readHistory = async (list) => Promise.all(list.map(async (version, index) => {
    const indexResult = await readSavedVersionRevisions(version.openId);
    return {
      version, entries: historyEntries(version, index, indexResult.records),
      changesRead: indexResult.status === "read", commentsRead: !version.commentsUnavailable,
    };
  }));
  const observed = await readHistory(commentVersions);
  for (let i = 1; i < 4; i++) reconcileHistoryVersions(observed[i - 1], observed[i], i);
  const commentEntries = observed.flatMap(version => version.entries);
  const old = commentEntries.find(entry => entry.foldedText.includes(searchKey("[原文]\nNEXT 📝")));
  assert.equal(old.version, 1);
  assert.equal(old.date, undefined);
  assert.deepEqual(historyMatch(old.text, searchKey("[原文]\nNEXT 📝")), { start: 5, end: old.text.length });
  const reply = commentEntries.find(entry => entry.kind === "reply");
  assert.equal(reply.version, 1);
  assert.deepEqual(reply.tags, ["根拠"]);
  const next = commentEntries.find(entry => entry.version === 2 && entry.resolved);
  assert.equal(next.absentAt, 3);
  assert.equal(old.absentAt, undefined, "edited text is still part of the same observed thread");
  assert.deepEqual(groupedHistoryRows(commentEntries).map(row => row.map(entry => entry.version)), [[1, 1, 2]]);
  const onlySource = groupedHistoryRows(commentEntries.filter(entry => entry.foldedText.includes("中間版だけ")));
  assert.deepEqual(onlySource, [[reply]]);
  assert.equal(lineageRangeText(commentVersions, { ...reply.ranges[0], version: reply.version }), "新案");
  assert.equal(commentVersions[4].commentsUnavailable, true, "a DOCX with comment marks but no comment part is incomplete");
  const missingAfter = await readHistory([commentVersions[1], commentVersions[4]]);
  reconcileHistoryVersions(missingAfter[0], missingAfter[1], 1);
  assert(missingAfter[0].entries.every(entry => entry.absentAt === undefined));
  const gapped = await readHistory([commentVersions[1], { ...commentVersions[2], gapBefore: true }]);
  reconcileHistoryVersions(gapped[0], gapped[1], 1);
  assert(gapped.every(version => version.entries.every(entry => entry.threadGroup === undefined)));
  const duplicate = await readHistory([commentVersions[1], {
    ...commentVersions[2], threads: [...commentVersions[2].threads, ...commentVersions[2].threads],
  }]);
  reconcileHistoryVersions(duplicate[0], duplicate[1], 1);
  assert(duplicate.every(version => version.entries.every(entry => entry.threadGroup === undefined && entry.absentAt === undefined)));
  const differentBranch = await readHistory([commentVersions[0], commentVersions[2], commentVersions[3]]);
  assert(differentBranch.flatMap(version => version.entries).every(entry => !entry.foldedText.includes("中間版だけ")));
  console.log("ok grouped durable threads, version-specific comments and replies, incomplete coverage and literal multiline search");

  const target = { driveId: "a", itemId: "b" };
  const fetches = [];
  const graphPage = "/v1.0/drives/a/items/b/versions";
  const paginated = async (url, options) => {
    fetches.push([String(url), options.headers]);
    return Response.json(url.search
      ? { value: [{ id: "1", lastModifiedDateTime: "2026-11-23T23:00:00Z" }] }
      : { value: [{ id: "2" }], "@odata.nextLink": `https://graph.microsoft.com${graphPage}?page=2` });
  };
  const signal = new AbortController().signal;
  assert.deepEqual((await listCloudVersionItems(target, "token", signal, paginated)).map(item => item.id), ["2", "1"]);
  assert.equal(fetches.length, 2);
  await assert.rejects(listCloudVersionItems(target, "token", signal, async () =>
    Response.json({ value: [], "@odata.nextLink": "https://elsewhere.example/" })), { reason: "invalid" });
  const headers = [];
  let requests = 0;
  const download = async (url, options) => {
    headers.push(options.headers?.Authorization);
    if (requests++ === 0) return new Response(null, { status: 429, headers: { "retry-after": "0" } });
    return url.hostname === "graph.microsoft.com"
      ? new Response(null, { status: 302, headers: { location: "https://download.example/version" } })
      : new Response(before);
  };
  assert.deepEqual(await getCloudVersionBytes(target, { id: "1" }, "token", signal, download), new Uint8Array(before));
  assert.deepEqual(headers, ["Bearer token", "Bearer token", undefined]);
  await assert.rejects(listCloudVersionItems(target, "expired", signal, async () =>
    new Response(null, { status: 401 })), { reason: "sign-in" });
  await assert.rejects(getCloudVersionBytes(target, { id: "expired" }, "token", signal, async () =>
    new Response(null, { status: 404 })), { reason: "missing" });

  const cache = join(folder, "cache");
  let downloads = 0;
  let account = "test-user";
  const client = {
    account: async () => ({ homeAccountId: account }),
    listVersions: async () => [{ id: "3" }, { id: "2" }, { id: "1" }],
    getVersionBytes: async (_target, item) => { downloads++; if (item.id === "3") return Buffer.from("invalid"); return before; },
  };
  const session = new CloudVersionImport(cache, { ...target, name: "File.docx", account, mtime: "2026-11-25T00:00:00Z", bytes: after }, client);
  await session.list();
  const imported = await session.import(["3", "1"], () => {});
  assert.equal(imported.versions[0].revisionCount, 4);
  const importedHistory = historyEntries(imported.versions[0], 0, (await readSavedVersionRevisions(imported.versions[0].openId)).records);
  assert(importedHistory.some(entry => entry.kind === "del" && entry.text === "旧案"));
  assert(importedHistory.some(entry => entry.kind === "comment" && entry.foldedText.includes("中間版のコメント")));
  assert.equal(imported.versions[1].gapBefore, true);
  assert.deepEqual(imported.failures, [{ id: "3", reason: "invalid" }]);
  const again = await session.import(["2", "1"], () => {});
  assert.equal(downloads, 3, "reuse a verified previously imported version");
  assert.equal(again.versions[0].sha256, again.versions[1].sha256);
  assert.notEqual(again.versions[0].versionId, again.versions[1].versionId);
  const saved = await readdir(join(cache, (await readdir(cache))[0]));
  assert.equal(saved.filter(name => name.endsWith(".docx")).length, 2, "one blob for identical past versions, plus the opened snapshot");
  assert.equal(saved.filter(name => name.endsWith(".json")).length, 2);
  account = "other";
  assert.equal((await readSavedVersionRevisions(again.versions[0].openId)).status, "unavailable");
  account = "test-user";
  const canceled = await session.import(["2", "1"], (done) => { if (done === 1) session.cancel(); });
  assert.equal(canceled.versions.length, 2);
  assert.deepEqual(canceled.failures, [{ id: "2", reason: "cancelled" }]);
  const retry = await session.import(["2", "1"], () => {});
  assert.equal(retry.failures.length, 0);
  assert.equal(downloads, 3, "resuming keeps the completed blobs");
  client.listVersions = async () => [{ id: "3" }];
  const retained = await session.list();
  assert.deepEqual(retained.items.filter(item => item.cachedOnly).map(item => item.id), ["2", "1"]);
  const expired = await session.import(["2", "1"], () => {});
  assert.equal(expired.versions[1].gapBefore, true, "cached-only entries preserve access without inferring continuity");
  assert.equal(downloads, 3);

  const sameTime = "2026-11-24T00:00:00Z";
  const openedDate = "2026-11-25T00:00:00Z";
  const refreshedList = async () => ["4", "3", "2", "1"].map(id => ({ id, modified: sameTime }));
  let authorized = false;
  const freshClient = {
    ...client,
    listVersions: async () => {
      if (!authorized) throw new CloudVersionError("sign-in");
      return refreshedList();
    },
    getVersionBytes: async (_target, item) => {
      downloads++;
      if (item.id === "3") throw new CloudVersionError("missing");
      return after;
    },
  };
  const restarted = new CloudVersionImport(cache, { ...target, name: "File.docx", account, mtime: openedDate, bytes: after }, freshClient);
  await assert.rejects(restarted.list(), { reason: "sign-in" });
  authorized = true;
  await restarted.list();
  const latest = await restarted.import(["4", "3", "2", "1"], () => {});
  assert.deepEqual(latest.failures, [{ id: "3", reason: "missing" }]);
  assert.deepEqual(latest.versions.map(version => version.label), ["File.docx · 1", "File.docx · 2", "File.docx · 3", "File.docx · 4", "File.docx"]);
  assert.equal(latest.versions[1].gapBefore, false, "Graph order is preserved for the same version date");
  assert.equal(latest.versions[1].sourceModified, sameTime);
  assert.equal(latest.versions[0].versionId, again.versions[0].versionId);
  assert.equal(latest.versions[0].fetchedAt, again.versions[0].fetchedAt, "restarting preserves successful imports");
  assert.deepEqual(latest.versions[0].paragraphs, latest.versions[3].paragraphs);
  assert.notEqual(latest.versions[0].sha256, latest.versions[3].sha256, "with identical visible text, track changes distinguish the bytes");
  assert(latest.versions[2].unavailable);
  assert.equal(latest.versions[2].openId, undefined);
  assert.equal(latest.versions.at(-1).gapBefore, true, "the open snapshot has no server position even with a later date and equal bytes");
  assert.notEqual(latest.versions.at(-1).versionId, latest.versions[3].versionId);
  assert.equal(latest.versions.at(-1).sha256, latest.versions[3].sha256);
  const reversed = await restarted.import(["1", "2", "3", "4"], () => {});
  assert.deepEqual(reversed.versions.map(version => [version.versionId, version.gapBefore]),
    latest.versions.map(version => [version.versionId, version.gapBefore]));
  const omitted = await restarted.import(["1", "4"], () => {});
  assert(omitted.versions[1].gapBefore);
  authorized = false;
  await assert.rejects(restarted.list(), { reason: "sign-in" });
  await assert.rejects(restarted.import(["1"], () => {}), { reason: "invalid" });

  if (process.env.DRA_HISTORY_BENCH) {
    const { performance } = await import("node:perf_hooks");
    const benchmarkFiles = [];
    for (let version = 0; version < 50; version++) {
      const file = join(folder, `benchmark-${version}.docx`);
      const body = Array.from({ length: 200 }, (_, i) =>
        para(node("調査結果 ") + ins(`version ${version}, measurement ${i}`, i + 1))).join("");
      await writeFile(file, await docx(body));
      benchmarkFiles.push(file);
    }
    const usage = () => process.memoryUsage().rss / 1024 ** 2;
    const initialRss = usage();
    let peakRss = initialRss;
    const started = performance.now();
    const benchmarkVersions = await readSelectedSavedVersions(benchmarkFiles, true);
    const records = [];
    let first;
    for (const [index, version] of benchmarkVersions.entries()) {
      const result = await readSavedVersionRevisions(version.openId);
      records.push(...historyEntries(version, index, result.records));
      first ??= performance.now() - started;
      peakRss = Math.max(peakRss, usage());
    }
    assert.equal(records.length, 10_000);
    const query = searchKey("version 49, measurement 199");
    const searchStarted = performance.now();
    const matches = records.filter(entry => entry.foldedText.includes(query));
    assert.equal(matches.length, 1);
    console.log(JSON.stringify({ versions: benchmarkVersions.length, records: records.length,
      firstBatchMs: Math.round(first), indexMs: Math.round(searchStarted - started),
      queryMs: Math.round((performance.now() - searchStarted) * 100) / 100,
      sampledRssIncreaseMiB: Math.round(peakRss - initialRss) }));
  }
  console.log("ok Graph pagination, retry, redirect auth, invalid DOCX, repeat import, distinct version provenance and account isolation");
} finally {
  await rm(folder, { recursive: true, force: true });
}
