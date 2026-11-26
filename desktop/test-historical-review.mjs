import assert from "node:assert/strict";
import { mkdtemp, writeFile, rm, readdir } from "node:fs/promises";
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
async function docx(body, comment) {
  const zip = new JSZip();
  zip.file("word/document.xml",
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}</w:body></w:document>`);
  if (comment) zip.file("word/comments.xml",
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="4" w:author="Reader">${para(node(comment))}</w:comment></w:comments>`);
  return zip.generateAsync({ type: "nodebuffer" });
}

try {
  const compiled = join(folder, "tests.cjs");
  await esbuild.build({
    stdin: { resolveDir: here, contents: `
      export { readSelectedSavedVersions, readSavedVersionRevisions } from "./src/savedLineage";
      export { historyEntries, historyMatch, searchKey } from "./src/historySearch";
      export { compareSavedVersions } from "../src/lib/lineage";
      export { projectHistoricalAnchor } from "../src/lib/historicalAnchors";
      export { listCloudVersionItems, getCloudVersionBytes, CloudVersionError } from "./src/cloudVersionGraph";
      export { CloudVersionImport } from "./src/cloudVersionImport";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile: compiled,
  });
  const {
    readSelectedSavedVersions, readSavedVersionRevisions, historyEntries, historyMatch, searchKey,
    compareSavedVersions, projectHistoricalAnchor, listCloudVersionItems, getCloudVersionBytes, CloudVersionImport, CloudVersionError,
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
  await writeFile(files[0], after);
  assert.equal((await readSavedVersionRevisions(versions[0].openId)).status, "changed");
  console.log("ok observed markup with identical visible text, duplicate IDs, exact source ranges, Unicode search and safe reconnection");

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
    const matches = records.filter(entry => searchKey(entry.text).includes(query));
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
