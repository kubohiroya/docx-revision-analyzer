import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const temp = await fs.mkdtemp(path.join(os.tmpdir(), "dra-version-branches-"));
const require = createRequire(import.meta.url);
const template = [
  "この資料は関係者間で参照するため、保存後にファイル名を記録してください。",
  "作成部署の連絡先は事務局で管理し、必要に応じて関係者へお知らせします。",
];
const root = [
  "教務部門は四月に運用方針を決定し、対象となる科目を各学部に通知しました。",
  "審査会議では登録上限を確認し、申請者に検討経過を説明することにしました。",
  "授業計画は出席者からの意見を集約して、月末までに個別に案内します。",
];
const a = [
  "野外観察の担当を自然科学部門に移し、地域別の調査票を採用します。",
  "標準試料の採取条件は海水温と潮流で分け、二人の記録を照合します。",
];
const b = [
  "閲覧端末の画面読み上げ機能を統合し、講義の録音も直接参照できます。",
  "映像配信の音声説明は五種類の言語を対象とし、確認会を別途開催します。",
];
const unrelated = [
  "家庭菜園では種の購入予定を組み替え、収穫時間の記録を残しておきます。",
  "職人が焼き上げた陶器の見本は、川向こうの展示室へ順番に運びました。",
];
const stamp = (hour) => new Date(Date.UTC(2026, 0, 1, hour)).toISOString();

function draft(id, hour, body, extra = {}) {
  return {
    id, label: id + ".docx", source: "selected",
    sourceModified: hour === undefined ? undefined : stamp(hour),
    paragraphs: body.map((text, index) => ({ index, text, kind: "body" })),
    trackedChanges: false, ...extra,
  };
}

try {
  const implementation = path.join(temp, "implementation.cjs");
  await esbuild.build({
    stdin: { resolveDir: here, contents: `
      export { inferVersionBranches } from "../src/lib/lineageBranches";
      export { readSelectedSavedVersions } from "./src/savedLineage";
    ` },
    bundle: true, platform: "node", format: "cjs", outfile: implementation,
  });
  const { inferVersionBranches, readSelectedSavedVersions } = require(implementation);
  const nodes = versions => new Map(inferVersionBranches(versions).map(node => [node.id, node]));
  const inferred = parentId => ({ status: "inferred", parentId });
  const held = reason => ({ status: "held", reason });

  const editions = [
    ["root", root], ["A1", [...root, ...a]], ["B1", [...root, ...b]],
    ["A2", [...root, ...a, "野外調査は午前中に実施し、担当者を当日の一覧で確認しました。"]],
    ["B2", [...root, ...b, "来年度は録音にも区切り位置を追加し、試行の結果を報告します。"]],
    ["unrelated", unrelated],
  ];
  const start = Math.ceil(Date.now() / 1000) * 1000 + 60_000;
  const files = [];
  for (const [index, [id, text]] of editions.entries()) {
    const zip = new JSZip();
    const body = [...template, ...text].map(sentence => `<w:p><w:r><w:t>${sentence}</w:t></w:r></w:p>`).join("");
    zip.file("word/document.xml", `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`);
    const file = path.join(temp, id + ".docx");
    await fs.writeFile(file, await zip.generateAsync({ type: "nodebuffer" }));
    const date = new Date(start + index * 1000);
    await fs.utimes(file, date, date);
    files.push(file);
  }
  const versions = (await readSelectedSavedVersions(files)).map(version => ({ ...version, id: path.parse(version.label).name }));
  const graph = inferVersionBranches(versions);
  const interleaved = nodes(versions);
  assert.deepEqual(graph, inferVersionBranches([...versions].reverse()));
  assert.deepEqual(JSON.parse(JSON.stringify(graph)), graph);
  assert.deepEqual(graph.map(node => node.id), editions.map(([id]) => id));
  assert.deepEqual(interleaved.get("A1").decision, inferred("root"));
  assert.deepEqual(interleaved.get("A2").decision, inferred("A1"));
  assert.deepEqual(interleaved.get("B2").decision, inferred("B1"));
  assert.deepEqual(interleaved.get("B1").decision, held("insufficient-distinctive-text"));
  assert.equal(interleaved.get("B1").candidates[0].parentId, "root");
  assert.deepEqual(interleaved.get("unrelated").decision, held("only-common-body-text"));
  assert.equal(interleaved.get("unrelated").candidates.length, 0);

  const withoutOutgroup = nodes(versions.slice(0, -1));
  assert.deepEqual(withoutOutgroup.get("A2").decision, inferred("A1"));
  assert.deepEqual(withoutOutgroup.get("B2").decision, inferred("B1"));
  assert.deepEqual(withoutOutgroup.get("A1").decision, held("only-common-body-text"));

  const linear = nodes([
    draft("L0", 0, [...root, ...a]), draft("L1", 1, [...root, ...a, ...b]),
    draft("L2", 2, [...root, ...b, ...unrelated]),
  ]);
  assert.deepEqual(linear.get("L1").decision, inferred("L0"));
  assert.deepEqual(linear.get("L2").decision, inferred("L1"));

  const candidate = interleaved.get("A2").candidates[0];
  assert.equal(candidate.distinctiveMatches, 2);
  assert.equal(candidate.order, "earlier");
  assert(candidate.sourceCoverage >= 0.5 && candidate.targetCoverage >= 0.3);
  const lookup = new Map(versions.map(version => [version.id, version]));
  function passage(versionId, range) {
    return lookup.get(versionId).paragraphs.find(paragraph => paragraph.index === range.paragraph)
      .text.slice(range.start, range.end);
  }
  assert.deepEqual(candidate.matches.filter(match => match.distinctive).map(match => {
    assert.equal(passage("A1", match.from), passage("A2", match.to));
    return passage("A2", match.to);
  }), a);
  for (const row of graph.filter(node => node.decision.status === "inferred")) {
    assert(row.time.selected.value > interleaved.get(row.decision.parentId).time.selected.value);
  }

  const previous = draft("previous", 1, a, { sha256: "a".repeat(64) });
  const copy = { ...previous, id: "copy", sourceModified: stamp(2), sha256: previous.sha256.toUpperCase() };
  const matches = draft("same-body", 3, a, { sha256: "b".repeat(64) });
  const repeated = nodes([matches, previous, copy]);
  assert.deepEqual(repeated.get("copy").decision, held("duplicate-file"));
  assert.equal(repeated.get("copy").duplicateOf, "previous");
  assert.deepEqual(repeated.get("same-body").decision, held("same-visible-text"));
  assert.deepEqual(repeated.get("same-body").sameVisibleText, ["previous"]);

  const sameClock = nodes([previous, draft("tied", 1, [...a, ...b], { sourceModified: "2026-01-01T10:00:00+09:00" })]);
  assert.deepEqual(sameClock.get("tied").decision, held("time-order-unconfirmed"));
  assert.equal(sameClock.get("tied").candidates[0].order, "same-time");
  assert.equal(sameClock.get("tied").time.selected.value, stamp(1));
  for (const [added, reason] of [
    [draft("missing", undefined, [...a, ...b]), "missing-time"],
    [draft("local-time", undefined, [...a, ...b], { sourceModified: "2026-01-01T02:00:00" }), "missing-time"],
    [draft("created", undefined, [...a, ...b], { fileCreatedAt: stamp(2) }), "creation-time-only"],
    [draft("conflicting", 2, [...a, ...b], { fileCreatedAt: stamp(3) }), "conflicting-times"],
    [draft("archived-too-early", 3, [...a, ...b], { source: "archive", archivedAt: stamp(2) }), "conflicting-times"],
  ]) {
    const row = nodes([previous, added]).get(added.id);
    assert.deepEqual(row.decision, held(reason));
    assert.equal(row.candidates[0].order, "unconfirmed-time");
    assert.equal(row.candidates[0].matches.length, 2);
  }

  const archiveA = draft("archive-a", 0, a, { source: "archive", archivedAt: stamp(1), number: 3, archiveId: "watch" });
  const archiveB = draft("archive-b", 1, [...a, ...b], { source: "archive", archivedAt: stamp(2), number: 2, archiveId: "watch" });
  const reversed = nodes([archiveA, archiveB]).get("archive-b");
  assert.deepEqual(reversed.time.selected, { source: "archivedAt", value: stamp(2) });
  assert.equal(reversed.time.archiveNumber, 2);
  assert.equal(reversed.candidates[0].order, "archive-order-conflict");
  assert.deepEqual(reversed.decision, held("time-order-unconfirmed"));
  assert.deepEqual(nodes([archiveA, { ...archiveB, archiveId: "other" }]).get("archive-b").decision, inferred("archive-a"));

  const competitors = nodes([previous, draft("B", 2, b), draft("both", 3, [...a, ...b])]).get("both");
  assert.deepEqual(competitors.decision, held("competing-candidates"));
  assert.deepEqual(competitors.candidates.map(row => row.distinctiveMatches), [2, 2]);
  const partial = nodes([draft("long", 1, [...a, ...b, ...root]), draft("partial", 2, [...a, ...unrelated])]).get("partial");
  assert.equal(partial.candidates[0].distinctiveMatches, 2);
  assert.deepEqual(partial.decision, held("insufficient-coverage"));
  const noMatch = draft("disconnected", 2, unrelated);
  assert.deepEqual(nodes([previous, noMatch]).get("disconnected").decision, held("no-shared-body-text"));
  assert.deepEqual(nodes([previous, { id: "unread", label: "unread.docx", source: "selected", unavailable: true }])
    .get("unread").decision, held("unavailable"));
  assert.throws(() => inferVersionBranches([previous, previous]), /Version IDs must be unique/);
  assert.deepEqual(inferVersionBranches([]), []);
  console.log("ok inferred interleaved branches from DOCX, evidence ranges, stable order, date uncertainty, duplicates and competing sources");
} finally {
  await fs.rm(temp, { recursive: true, force: true });
}
