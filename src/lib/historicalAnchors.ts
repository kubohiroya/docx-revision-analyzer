import type { LineageLink, LineageRange, SavedParagraphRange, SavedVersionLineage } from "./lineage";

export type HistoricalMatch =
  | { kind: "range" | "partial" | "boundary" | "related"; ranges: LineageRange[] }
  | { kind: "ambiguous"; ranges: LineageRange[] }
  | { kind: "deleted"; ranges: [] }
  | { kind: "unknown"; ranges: []; reason: "no-source" | "no-path" | "gap" };

function sameParagraph(a: SavedParagraphRange, b: SavedParagraphRange): boolean {
  return a.paragraph === b.paragraph;
}

function touches(a: SavedParagraphRange, b: SavedParagraphRange): boolean {
  return sameParagraph(a, b) && (a.start === a.end
    ? b.start <= a.start && a.start <= b.end
    : a.start < b.end && b.start < a.end);
}

function mapSpan(range: LineageRange, links: readonly LineageLink[]) {
  const result: LineageRange[] = [];
  let covered = 0;
  for (const link of links) {
    for (const { from, to } of link.unchanged) {
      if (!sameParagraph(range, from)) continue;
      const start = Math.max(range.start, from.start);
      const end = Math.min(range.end, from.end);
      if (start >= end) continue;
      result.push({ ...to, start: to.start + start - from.start, end: to.start + end - from.start });
      covered += end - start;
    }
  }
  return { result, covered };
}

function mapBoundary(range: LineageRange, links: readonly LineageLink[]): LineageRange | undefined {
  const pairs = links.flatMap(link => link.unchanged)
    .filter(pair => sameParagraph(pair.from, range));
  const left = pairs.filter(({ from }) => from.start < range.start && from.end >= range.start);
  const right = pairs.filter(({ from }) => from.start <= range.end && from.end > range.end);
  if (left.length !== 1 || right.length !== 1) return undefined;
  const a = left[0], b = right[0];
  const start = a.to.start + range.start - a.from.start;
  const end = b.to.start + range.end - b.from.start;
  return sameParagraph(a.to, b.to) && start === end ? { ...a.to, start, end } : undefined;
}

function distinctRanges(ranges: readonly LineageRange[]): LineageRange[] {
  return [...new Map(ranges.map(range => [`${range.version}:${range.paragraph}:${range.start}:${range.end}`, range])).values()];
}

/** Follow an ordered, explicitly scoped path using the actual aligned character runs. */
export function projectHistoricalAnchor(
  lineage: SavedVersionLineage, origin: number, target: number, source: readonly SavedParagraphRange[],
): HistoricalMatch {
  if (!source.length) return { kind: "unknown", ranges: [], reason: "no-source" };
  if (origin > target || origin < 0 || target >= lineage.versions.length) return { kind: "unknown", ranges: [], reason: "no-path" };
  const route = lineage.comparisons.slice(origin, target);
  if (route.some(step => step.unavailable || step.recordingGap)) return { kind: "unknown", ranges: [], reason: "gap" };
  let ranges = source.map(range => ({ ...range, version: origin }));
  let related = false;
  let partial = false;
  for (const step of route) {
    const links = step.paragraph;
    const next: LineageRange[] = [];
    for (const range of ranges) {
      const possible = links.filter(link => link.from.some(from => touches(range, from)));
      if (possible.some(link => link.kind === "ambiguous" || link.kind === "copied" || link.kind === "reappeared")) {
        const candidates = possible.flatMap(link => link.to);
        return { kind: "ambiguous", ranges: distinctRanges(candidates) };
      }
      const matches = possible.filter(link => link.to.length && link.from.every(from => from.version === range.version));
      if (!related) {
        const { result, covered } = mapSpan(range, matches);
        if (result.length) {
          partial ||= covered < range.end - range.start;
          next.push(...result);
          continue;
        }
        const boundary = mapBoundary(range, matches);
        if (boundary) {
          partial ||= range.start !== range.end;
          next.push(boundary);
          continue;
        }
      }
      if (matches.length && (related || range.start !== range.end)) {
        related = true;
        next.push(...matches.flatMap(link => link.to));
      } else if (matches.length) {
        return { kind: "unknown", ranges: [], reason: "no-path" };
      } else if (ranges.length > 1) {
        partial = true;
      }
    }
    if (!next.length) return { kind: "deleted", ranges: [] };
    ranges = distinctRanges(next);
  }
  const kind = related ? "related" : ranges.every(range => range.start === range.end) ? "boundary" : partial ? "partial" : "range";
  return { kind, ranges };
}
