import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";
import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const here = path.dirname(fileURLToPath(import.meta.url));
const temp = await fs.mkdtemp(path.join(os.tmpdir(), "dra-saved-versions-"));
const require = createRequire(import.meta.url);

async function docx(paragraphs) {
  const zip = new JSZip();
  const text = paragraphs.map((p) => typeof p === "string" ? { text: p } : p);
  const paragraphsXml = text.map(({ text, id }) => `<w:p${id ? ` w14:paraId="${id}"` : ""}><w:r><w:t>${text}</w:t></w:r></w:p>`).join("");
  zip.file("word/document.xml", `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphsXml}</w:body></w:document>`);
  return zip.generateAsync({ type: "nodebuffer" });
}

try {
  const file = path.join(temp, "implementation.cjs");
  await esbuild.build({
    stdin: { resolveDir: here, contents: `
      export { compareSavedVersions, lineageRangeText, traceLineage, readSavedTextVersion } from "../src/lib/lineage";
      export { WatchArchive } from "../src/node/docWatcher";
      export { readSelectedSavedVersions, readSavedVersionArchive } from "./src/savedLineage";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile: file,
  });
  const { compareSavedVersions, lineageRangeText, traceLineage, readSavedTextVersion,
    WatchArchive, readSelectedSavedVersions, readSavedVersionArchive } = require(file);
  const makeVersions = async (drafts) => Promise.all(drafts.map(async (paragraphs, index) =>
    readSavedTextVersion(await docx(paragraphs), { label: `draft-${index + 1}.docx`, source: "selected" })));
  const words = (value, ranges) => ranges.map((range) => lineageRangeText(value.versions, range));

  const initial = "制度によって欠席率が20％低下した。この結果は別の調査でも同じだった。";
  const revised = "制度の導入後、欠席率が20％低下した。この結果は別の調査でも同じだった。";
  const tracked = (text) => ({ text, id: "01020304" });
  const first = "The evaluation used three independent tests.";
  const second = "We archived the raw observations for future readers.";
  const anchor = "すべての調査票は📝別のフォルダに保管した。";
  const versions = await makeVersions([
    [tracked(initial), anchor, `${first} ${second}`],
    [anchor, tracked(revised), `${first} ${second}`],
    [anchor, tracked(revised), first, second],
    [anchor, tracked(revised), `${first} ${second}`],
  ]);
  assert(versions.every((v) => !v.trackedChanges && !v.archivedAt), "local ordering works without edit or save times");
  const result = compareSavedVersions(versions);
  const move = result.comparisons[0].paragraph.find((link) => link.kind === "rewritten");
  assert(move?.reordered, "a moved and rewritten paragraph keeps its relationship");
  assert(move.evidence.includes("paragraphId"));
  assert.deepEqual(words(result, move.from), [initial]);
  assert.deepEqual(words(result, move.to), [revised]);
  assert(move.unchanged.some(({ to }) => lineageRangeText(result.versions, to).includes("欠席率が20％低下した")));
  const sentence = result.comparisons[0].sentence.find((link) => link.kind === "rewritten" && link.reordered);
  assert.deepEqual(words(result, sentence.from), ["制度によって欠席率が20％低下した。"]);
  assert.deepEqual(words(result, sentence.to), ["制度の導入後、欠席率が20％低下した。"]);
  for (const comparison of result.comparisons) {
    for (const link of [...comparison.paragraph, ...comparison.sentence]) {
      for (const { from, to } of link.unchanged) {
        assert.equal(lineageRangeText(result.versions, from), lineageRangeText(result.versions, to));
      }
    }
  }
  const split = result.comparisons[1].paragraph.find((link) => link.kind === "split");
  assert.deepEqual(words(result, split.to), [first, second]);
  const merged = result.comparisons[2].paragraph.find((link) => link.kind === "merged");
  assert.deepEqual(words(result, merged.from), [first, second]);
  assert(traceLineage(result, "paragraph", split.to[1]).includes(merged));
  assert(traceLineage(result, "paragraph", move.to[0]).some((link) => link.to.some((range) => range.version === 3)));

  const withWords = compareSavedVersions(versions, { japaneseWordHints: true });
  const japaneseMove = withWords.comparisons[0].paragraph.find((link) => link.kind === "rewritten");
  assert.deepEqual(japaneseMove.to, move.to, "Japanese word segmentation preserves the character baseline on the labeled rewrite");
  assert(japaneseMove.evidence.includes("japaneseWords"));
  assert.deepEqual(japaneseMove.unchanged, move.unchanged);
  console.log("ok Japanese character and word-hint comparison, English split/merge, sentence and paragraph movement");

  const phrase = "A paragraph that was present before the review.";
  const left = "Left anchor for the original text.";
  const right = "Right anchor for the original text.";
  const history = compareSavedVersions(await makeVersions([
    [left, phrase, right], ["New first paragraph.", left, phrase, right],
    [left, phrase, right, phrase], [left, right], [left, right, phrase],
  ]));
  assert(history.comparisons[0].paragraph.every((link) => !link.reordered), "insertion shifts indices without moving existing text");
  const copy = history.comparisons[1].paragraph.find((link) => link.kind === "copied");
  assert.equal(copy.to[0].paragraph, 3);
  assert.equal(copy.from[0].paragraph, 2);
  assert(history.comparisons[2].paragraph.some((link) => link.kind === "removed"));
  const uncertain = history.comparisons[3].paragraph.find((link) => link.kind === "ambiguous");
  assert.equal(uncertain.from.length, 2, "two previous copies must remain separate candidates");
  assert.notEqual(uncertain.from[0].paragraph, uncertain.from[1].paragraph);
  const comesBack = compareSavedVersions(await makeVersions([[left, phrase], [left], [left, phrase]]));
  const returnLink = comesBack.comparisons[1].paragraph.find((link) => link.kind === "reappeared");
  assert.equal(returnLink.from[0].version, 0);
  assert(comesBack.comparisons[0].paragraph.some((link) => link.kind === "removed"));
  const duplicates = compareSavedVersions(await makeVersions([[phrase, phrase], [phrase, phrase]]));
  assert(duplicates.comparisons[0].paragraph.every((link) => link.kind === "ambiguous" && link.from.length === 2));
  assert.equal(traceLineage(duplicates, "paragraph", duplicates.comparisons[0].paragraph[0].to[0]).length, 1);
  const similar = compareSavedVersions(await makeVersions([
    ["The north district reported several incidents during the review period.",
      "The south district reported several incidents during the review period."],
    ["The coast district reported several incidents during the review period."],
  ]));
  assert(similar.comparisons[0].paragraph.some((link) => link.kind === "ambiguous" && link.from.length === 2));
  const multipleOrigins = compareSavedVersions(await makeVersions([
    [{ text: phrase, id: "01020304" }, { text: phrase, id: "02030405" }],
    [{ text: phrase, id: "01020304" }, { text: phrase, id: "02030405" }, phrase],
  ]));
  assert(multipleOrigins.comparisons[0].paragraph.some((link) => link.kind === "ambiguous" && link.from.length === 2));
  const boilerplate = compareSavedVersions(await makeVersions([Array(20).fill(phrase), [phrase]]));
  assert.equal(boilerplate.comparisons[0].paragraph[0].candidateCount, 20);
  console.log("ok copy without identity collapse, relative order, reappearance and repeated text ambiguity");

  const original = path.join(temp, "original.docx");
  const drafts = [];
  const archive = WatchArchive.open(path.join(temp, "watch"), original);
  for (const [index, paragraphs] of [[left, phrase], [left, phrase, right], [left, right]].entries()) {
    const bytes = await docx(paragraphs);
    const name = path.join(temp, `file-${index + 1}.docx`);
    drafts.push(name);
    await fs.writeFile(name, bytes);
    const date = new Date(Date.UTC(2026, 0, 1 + index));
    await archive.addVersion(bytes, date, date);
  }
  const manifest = path.join(archive.dir, "snapshots.json");
  const savedManifest = await fs.readFile(manifest);
  const selected = await readSelectedSavedVersions([drafts[2], drafts[0], drafts[1]]);
  assert.deepEqual(selected.map((v) => v.label), ["file-3.docx", "file-1.docx", "file-2.docx"]);
  const recorded = await readSavedVersionArchive(archive);
  assert.deepEqual(recorded.map((v) => v.number), [1, 2, 3]);
  assert.equal(recorded[0].sha256, selected[1].sha256);
  assert.equal(archive.versions()[0].archivedAt, recorded[0].archivedAt);
  for (const [index, v] of recorded.entries()) {
    assert.deepEqual(v.paragraphs, selected[(index + 1) % 3].paragraphs);
  }
  await fs.unlink(path.join(archive.dir, archive.versions()[1].file));
  const missing = await readSavedVersionArchive(archive);
  assert(missing[1].unavailable);
  assert(compareSavedVersions(missing).comparisons.every((p) => p.unavailable && p.paragraph.length === 0));
  assert.deepEqual(await fs.readFile(manifest), savedManifest);
  const gap = compareSavedVersions([comesBack.versions[0], comesBack.versions[1], { ...comesBack.versions[2], gapBefore: true }]);
  assert(gap.comparisons[1].recordingGap);
  assert(!gap.comparisons[1].paragraph.some((link) => link.kind === "reappeared"));
  console.log("ok explicit file order, verified archived order, missing saved version and recording gaps");
} finally {
  await fs.rm(temp, { recursive: true, force: true });
}
