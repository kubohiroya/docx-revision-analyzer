// Read-only check of the three ordered DOCX versions prepared for issue #111.
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";

const files = process.argv.slice(2).map(file => resolve(file));
if (files.length !== 3 || new Set(files).size !== 3) {
  console.error("Usage: node desktop/scripts/verify-history-case.mjs 01-A.docx 02-B.docx 03-C.docx");
  process.exit(2);
}
const here = dirname(fileURLToPath(import.meta.url));
const digest = async file => createHash("sha256").update(await readFile(file)).digest("hex");
const before = await Promise.all(files.map(digest));
const temporary = await mkdtemp(join(tmpdir(), "dra-verify-history-"));
try {
  const outfile = join(temporary, "check.cjs");
  await esbuild.build({
    stdin: { resolveDir: here, contents: `
      export { readSelectedSavedVersions, readSavedVersionRevisions } from "../src/savedLineage";
      export { historyEntries } from "../src/historySearch";
      export { compareSavedVersions, lineageRangeText } from "../../src/lib/lineage";
      export { projectHistoricalAnchor } from "../../src/lib/historicalAnchors";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile,
  });
  const {
    readSelectedSavedVersions, readSavedVersionRevisions, historyEntries,
    compareSavedVersions, lineageRangeText, projectHistoricalAnchor,
  } = createRequire(import.meta.url)(outfile);
  const versions = await readSelectedSavedVersions(files, true);
  assert(versions.every(v => !v.unavailable && v.openId), "All three versions must be readable");
  assert.equal(new Set(versions.map(v => v.versionId)).size, 3);
  const history = await Promise.all(versions.map(async (v, i) => {
    const revisions = await readSavedVersionRevisions(v.openId);
    assert.equal(revisions.status, "read", basename(files[i]));
    assert(!v.commentsUnavailable, `Comments unavailable in ${basename(files[i])}`);
    return historyEntries(v, i, revisions.records);
  }));
  const markers = [["del", "赤鉛筆111"], ["comment", "根拠確認111"], ["reply", "返信確認111"]];
  const matches = Object.fromEntries(markers.map(([kind, text]) => {
    const versionIndexes = history.flatMap((rows, i) => rows.some(row => row.kind === kind && row.text.includes(text)) ? [i] : []);
    assert.deepEqual(versionIndexes, [1], `${kind} must be observed only in edition B`);
    return [kind, history[1].find(row => row.kind === kind && row.text.includes(text))];
  }));
  assert(matches.comment.ranges.length, "The source comment must have a range in B");
  const lineage = compareSavedVersions(versions);
  const destination = projectHistoricalAnchor(lineage, 1, 2, matches.comment.ranges);
  assert.equal(destination.kind, "range", "The source comment must map to a unique range in C");
  assert(destination.ranges.some(range => lineageRangeText(lineage.versions, range)?.includes("K111")),
    "The destination range must contain the expected text");
  console.log(JSON.stringify({ source: basename(files[1]), destination: basename(files[2]),
    matchedKinds: Object.keys(matches), destinationKind: destination.kind, inputSha256: before }, null, 2));
} finally {
  await rm(temporary, { recursive: true, force: true });
  assert.deepEqual(await Promise.all(files.map(digest)), before, "Input documents must not change");
}
