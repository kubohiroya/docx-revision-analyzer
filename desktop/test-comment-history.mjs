import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "dra-comment-history-"));
const require = createRequire(import.meta.url);
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const quote = "欠席率が20％低下";
const prose = "#出典確認\n報告書: https://example.org/report.pdf#page=12";
const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function docx({ resolved = false, position = 1, missing = false, noAnchor = false, identical = false, oldReply = false } = {}) {
  const zip = new JSZip();
  const placement = (n) => {
    const anchored = !missing && !noAnchor && position === n;
    const prefix = identical ? "同じ前置き".repeat(8) : `段落 ${n} `;
    const suffix = identical ? "同じ後続文".repeat(8) : "。";
    const start = anchored ? '<w:commentRangeStart w:id="0"/>' : "";
    const end = anchored ? '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>' : "";
    const id = identical ? "" : ` w14:paraId="0000000${n}"`;
    return `<w:p${id}>${run(prefix)}${start}${run(quote)}${end}${run(suffix)}</w:p>`;
  };
  zip.file("word/document.xml", `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${placement(1)}${placement(2)}</w:body></w:document>`);
  if (!missing) {
    zip.file("word/comments.xml", `<w:comments xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}">
      <w:comment w:id="0" w:author="Writer"><w:p w14:paraId="00000010">${run("#出典確認")}<w:r><w:br/></w:r>${run("報告書: https://example.org/report.pdf#page=12")}</w:p></w:comment>
      <w:comment w:id="1" w:author="Reviewer"><w:p w14:paraId="00000011">${run("#判断 ")}<w:hyperlink r:id="rId1">${run("表3")}</w:hyperlink></w:p></w:comment>
      ${oldReply ? '<w:comment w:id="2" w:author="Reviewer"><w:p w14:paraId="00000012"><w:r><w:t>以前の返信</w:t></w:r></w:p></w:comment>' : ""}
    </w:comments>`);
    zip.file("word/commentsExtended.xml", `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
      <w15:commentEx w15:paraId="00000010" w15:done="${Number(resolved)}"/>
      <w15:commentEx w15:paraId="00000011" w15:paraIdParent="00000010"/>
      <w15:commentEx w15:paraId="00000012" w15:paraIdParent="00000010"/>
    </w15:commentsEx>`);
    zip.file("word/_rels/comments.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="${R}/hyperlink" TargetMode="External" Target="https://example.org/source?a=1&amp;b=2#table3"/>
    </Relationships>`);
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

try {
  const bundle = path.join(tmp, "test.cjs");
  await esbuild.build({
    stdin: { contents: `
      export { WatchArchive } from "../src/node/docWatcher";
      export { diffComments, extractComments, textHash } from "../src/lib/docxComments";
      export { commentReferences, commentLinkUrl } from "../src/lib/commentReferences";
      export { listCommentHistory, readCommentHistory } from "./src/commentHistory";
    `, resolveDir: here },
    bundle: true, platform: "node", format: "cjs", outfile: bundle, logLevel: "silent",
  });
  const { WatchArchive, diffComments, extractComments, textHash, commentReferences, commentLinkUrl, listCommentHistory, readCommentHistory } = require(bundle);
  const root = path.join(tmp, "watch");
  const source = path.join(tmp, "report.docx");
  assert.equal(WatchArchive.find(root, source), undefined);
  await assert.rejects(fs.stat(root), { code: "ENOENT" });
  const archive = WatchArchive.open(root, source);
  let clock = Date.parse("2026-06-01T10:00:00Z");
  const capture = async (a, options) => {
    const now = new Date(clock += 60_000);
    const bytes = await docx(options);
    const version = await a.addVersion(bytes, now, now);
    const events = a.appendEvents(diffComments(a.commentState(), await extractComments(bytes), now, version.sha256));
    return { version, events };
  };
  const first = await capture(archive, { oldReply: true });
  const key = first.events[0].key;
  assert.equal(key, `h:${textHash(`Writer\0\0${prose.replace("\n", "")}`)}`, "legacy text-derived keys survive line break extraction");
  const second = await capture(archive, { resolved: true });
  const third = await capture(archive, { resolved: true, position: 2 });
  assert.deepEqual(third.events, [], "an unchanged comment can move without producing a new event");
  const fourth = await capture(archive, { missing: true });
  const manifest = path.join(archive.dir, "snapshots.json");
  const log = path.join(archive.dir, "comments.jsonl");
  const before = await Promise.all([fs.readFile(manifest), fs.readFile(log), fs.stat(path.join(archive.dir, "watch.json"))]);

  const reopened = WatchArchive.find(root, source);
  assert(reopened);
  const rows = listCommentHistory(reopened);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].present, false);
  assert.deepEqual(rows[0].tags, ["出典確認", "判断"]);
  const detail = await readCommentHistory(reopened, key);
  assert.equal(detail.version.n, third.version.n);
  assert.deepEqual(detail.passage, { prefix: "段落 2 ", quote, suffix: "。" });
  assert.equal(detail.comments.length, 2, "a reply from an older version is not mixed into the selected version");
  assert.deepEqual(detail.observations.map((e) => [e.type, e.version]), [["resolved", second.version.n], ["removed", fourth.version.n]]);
  assert.deepEqual(commentReferences(detail.comments[0].text), {
    tags: ["出典確認"], links: [{ label: "https://example.org/report.pdf#page=12", url: "https://example.org/report.pdf#page=12" }],
  });
  assert.deepEqual(commentReferences(detail.comments[1].text, detail.comments[1].links), {
    tags: ["判断"], links: [{ label: "表3", url: "https://example.org/source?a=1&b=2#table3" }],
  });
  assert.equal(commentLinkUrl("file:///tmp/report.pdf"), undefined);
  assert.equal(commentLinkUrl("javascript:alert(1)"), undefined);
  assert.deepEqual(await fs.readFile(manifest), before[0]);
  assert.deepEqual(await fs.readFile(log), before[1]);
  assert.equal((await fs.stat(path.join(archive.dir, "watch.json"))).mtimeMs, before[2].mtimeMs);

  await fs.unlink(path.join(archive.dir, third.version.file));
  const partial = await readCommentHistory(reopened, key);
  assert.equal(partial.version.n, second.version.n);
  assert.equal(partial.versionUnavailable, true);
  await Promise.all([first, second].map(({ version }) => fs.unlink(path.join(archive.dir, version.file))));
  const fallback = await readCommentHistory(reopened, key);
  assert.equal(fallback.version, undefined);
  assert.equal(fallback.versionUnavailable, true);
  assert.equal(fallback.recordedAnchor.quote, quote);
  assert.equal(fallback.comments.length, 2, "fallback uses the recorded thread before disappearance");
  await fs.writeFile(manifest, "{");
  assert.equal((await readCommentHistory(reopened, key)).recordedAnchor.quote, quote);

  const ambiguous = WatchArchive.open(root, path.join(tmp, "ambiguous.docx"));
  await capture(ambiguous, { identical: true });
  await capture(ambiguous, { identical: true, noAnchor: true });
  const uncertain = await readCommentHistory(ambiguous, key);
  assert.equal(uncertain.version.n, 2);
  assert.equal(uncertain.passage, undefined, "a duplicate quote without a matching context must not highlight the wrong paragraph");
  assert.equal(uncertain.recordedAnchor.quote, quote);

  const contextual = WatchArchive.open(root, path.join(tmp, "contextual.docx"));
  await capture(contextual, {});
  await capture(contextual, { noAnchor: true });
  const matched = await readCommentHistory(contextual, key);
  assert.equal(matched.locatedByContext, true);
  assert.deepEqual(matched.passage, { prefix: "段落 1 ", quote, suffix: "。" });
  console.log("ok saved comment threads, archived passages, replies, tags, links, observations, and missing data");
} finally {
  await fs.rm(tmp, { recursive: true, force: true });
}
