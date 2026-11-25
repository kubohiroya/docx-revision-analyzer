import { parseDocxLayout, type LayoutCommentMark, type Para } from "./docxLayout";
import type { DocxInput } from "./input";
import { compareUnits, prepareUnits, type Unit } from "./lineageMatching";

export type LineageLevel = "paragraph" | "sentence";
export type LineageKind = "retained" | "rewritten" | "split" | "merged" | "copied" | "reappeared" | "added" | "removed" | "ambiguous";
export type MatchEvidence = "exact" | "paragraphId" | "neighbors" | "heading" | "characterOverlap" | "japaneseWords" | "joined" | "past";

export interface VersionIdentity {
  label: string;
  source: "selected" | "archive";
  number?: number;
  sha256?: string;
  archivedAt?: string;
  sourceModified?: string;
  gapBefore?: boolean;
}

export interface VersionParagraph {
  index: number;
  text: string;
  kind: "body" | "heading";
  paraId?: string;
}

export type SavedTextVersion = VersionIdentity & (
  | { paragraphs: VersionParagraph[]; trackedChanges: boolean; unavailable?: false }
  | { unavailable: true }
);

/** Offsets are UTF-16 positions within the original paragraph text for this version. */
export interface SavedParagraphRange {
  paragraph: number;
  start: number;
  end: number;
}

export interface LineageRange extends SavedParagraphRange {
  version: number;
}

export interface LineageLink {
  kind: LineageKind;
  from: LineageRange[];
  to: LineageRange[];
  evidence: MatchEvidence[];
  candidateCount?: number;
  reordered?: boolean;
  alignmentPartial?: boolean;
  unchanged: { from: LineageRange; to: LineageRange }[];
}

export interface VersionComparison {
  before: number;
  after: number;
  unavailable: boolean;
  recordingGap: boolean;
  paragraph: LineageLink[];
  sentence: LineageLink[];
}

export interface SavedVersionLineage {
  versions: readonly SavedTextVersion[];
  comparisons: VersionComparison[];
}

export interface LineageOptions {
  /** Optional word boundary hints for Japanese. Exact ranges always use original characters. */
  japaneseWordHints?: boolean;
}

export async function readSavedTextVersion(
  input: DocxInput, identity: VersionIdentity, commentRanges?: Map<string, SavedParagraphRange[]>,
): Promise<SavedTextVersion> {
  const rawMarks = commentRanges && new Map<number, LayoutCommentMark[]>();
  const model = await parseDocxLayout(input, rawMarks);
  const paragraphs: VersionParagraph[] = [];
  const marks = new Map<string, Array<{ kind: LayoutCommentMark["kind"]; paragraph: number; offset: number }>>();
  let trackedChanges = false;
  for (const p of model.paras) {
    const pending = rawMarks?.get(p.index);
    if (p.kind === "table" && !pending) continue;
    const { text, offsets } = visibleText(p, pending);
    pending?.forEach((mark, index) => {
      const existing = marks.get(mark.id) ?? [];
      existing.push({ kind: mark.kind, paragraph: p.index, offset: offsets[index] });
      marks.set(mark.id, existing);
    });
    if (p.kind === "table") continue;
    trackedChanges ||= Boolean(p.markIns || p.markDel || p.segs.some((s) => s.ins || s.del));
    if (text.trim()) paragraphs.push({ index: p.index, kind: p.kind, text, paraId: p.paraId });
  }
  if (commentRanges) {
    const byIndex = new Map(paragraphs.map((p) => [p.index, p]));
    for (const [id, entries] of marks) {
      const starts = entries.filter((mark) => mark.kind === "commentRangeStart");
      const ends = entries.filter((mark) => mark.kind === "commentRangeEnd");
      const refs = entries.filter((mark) => mark.kind === "commentReference");
      if (!starts.length && !ends.length && refs.length === 1 && byIndex.has(refs[0].paragraph)) {
        const { paragraph, offset } = refs[0];
        commentRanges.set(id, [{ paragraph, start: offset, end: offset }]);
      } else if (starts.length === 1 && ends.length === 1) {
        const start = starts[0];
        const end = ends[0];
        if (!byIndex.has(start.paragraph) || !byIndex.has(end.paragraph)) continue;
        const ranges: SavedParagraphRange[] = [];
        for (let index = start.paragraph; index <= end.paragraph; index++) {
          const p = byIndex.get(index);
          if (!p) continue;
          const from = index === start.paragraph ? start.offset : 0;
          const to = index === end.paragraph ? end.offset : p.text.length;
          if (from < to) ranges.push({ paragraph: index, start: from, end: to });
        }
        if (ranges.length) commentRanges.set(id, ranges);
      }
    }
  }
  return { ...identity, paragraphs, trackedChanges };
}

/** Use the same visible text and offsets for comparison ranges and Word comment marks. */
function visibleText(p: Para, marks: LayoutCommentMark[] = []): { text: string; offsets: number[] } {
  let text = "";
  const offsets: number[] = [];
  for (let i = 0; i <= p.segs.length; i++) {
    while (marks[offsets.length]?.segment === i) offsets.push(text.length);
    const seg = p.segs[i];
    if (!seg || seg.del) continue;
    if (seg.kind === "text") text += seg.text;
    else if (seg.kind === "lineBreak" || seg.kind === "pageBreak") text += "\n";
  }
  return { text, offsets };
}

export function compareSavedVersions(versions: readonly SavedTextVersion[], options: LineageOptions = {}): SavedVersionLineage {
  const sentenceSegmenter = new Intl.Segmenter("ja", { granularity: "sentence" });
  const wordSegmenter = options.japaneseWordHints ? new Intl.Segmenter("ja", { granularity: "word" }) : undefined;
  const levels: Record<LineageLevel, Unit[][]> = { paragraph: [], sentence: [] };
  versions.forEach((version, index) => {
    const paragraphs = version.unavailable ? [] : version.paragraphs;
    levels.paragraph.push(prepareUnits(paragraphs, index, undefined, wordSegmenter));
    levels.sentence.push(prepareUnits(paragraphs, index, sentenceSegmenter, wordSegmenter));
  });
  const comparisons: VersionComparison[] = [];
  const past: Record<LineageLevel, Map<string, Unit[]>> = { paragraph: new Map(), sentence: new Map() };
  for (let after = 1; after < versions.length; after++) {
    const before = after - 1;
    const unavailable = Boolean(versions[before].unavailable || versions[after].unavailable);
    const recordingGap = Boolean(versions[after].gapBefore);
    const compare = (level: LineageLevel) => {
      if (unavailable || recordingGap) past[level].clear();
      if (unavailable) return [];
      const left = levels[level][before];
      const links = compareUnits(left, levels[level][after], past[level]);
      if (!recordingGap) {
        const occurrences = new Map<string, Unit[]>();
        for (const unit of left) {
          const list = occurrences.get(unit.key) ?? [];
          list.push(unit);
          occurrences.set(unit.key, list);
        }
        for (const [key, units] of occurrences) past[level].set(key, units);
      }
      return links;
    };
    comparisons.push({
      before, after, unavailable, recordingGap,
      paragraph: compare("paragraph"),
      sentence: compare("sentence"),
    });
  }
  return { versions, comparisons };
}

export function lineageRangeText(versions: readonly SavedTextVersion[], range: LineageRange): string {
  const version = versions[range.version];
  return version && !version.unavailable
    ? version.paragraphs.find((p) => p.index === range.paragraph)?.text.slice(range.start, range.end) ?? ""
    : "";
}

function rangeKey(range: LineageRange): string {
  return `${range.version}:${range.paragraph}:${range.start}:${range.end}`;
}

/** Walk confirmed relationships in both directions from one occurrence. */
export function traceLineage(lineage: SavedVersionLineage, level: LineageLevel, occurrence: LineageRange): LineageLink[] {
  const all = lineage.comparisons.flatMap((comparison) => comparison[level]);
  const byRange = new Map<string, LineageLink[]>();
  for (const link of all) {
    for (const range of [...link.from, ...link.to]) {
      const key = rangeKey(range);
      const neighbors = byRange.get(key) ?? [];
      neighbors.push(link);
      byRange.set(key, neighbors);
    }
  }
  const visitedRanges = new Set<string>();
  const visitedLinks = new Set<LineageLink>();
  const queue = [occurrence];
  for (let i = 0; i < queue.length; i++) {
    const key = rangeKey(queue[i]);
    if (visitedRanges.has(key)) continue;
    visitedRanges.add(key);
    for (const link of byRange.get(key) ?? []) {
      if (visitedLinks.has(link)) continue;
      visitedLinks.add(link);
      if (link.kind !== "ambiguous") queue.push(...link.from, ...link.to);
    }
  }
  return all.filter((link) => visitedLinks.has(link));
}
