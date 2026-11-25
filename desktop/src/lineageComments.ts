import type { LineageLevel, LineageLink, LineageRange, SavedParagraphRange, SavedVersionLineage } from "../../src/lib/lineage";
import type { SavedLineageVersion, SavedVersionThread } from "./shared";

export interface VersionThreadOccurrence {
  version: number;
  thread: SavedVersionThread;
}

function rangeKey(range: LineageRange): string {
  return `${range.version}:${range.paragraph}:${range.start}:${range.end}`;
}

function intersects(mark: SavedParagraphRange, text: LineageRange, paragraphLength: number): boolean {
  if (mark.paragraph !== text.paragraph) return false;
  if (mark.start === mark.end) {
    return text.start <= mark.start && (mark.start < text.end || (mark.start === text.end && text.end === paragraphLength));
  }
  return mark.start < text.end && text.start < mark.end;
}

export class LineageComments {
  private threadIndex = new Map<string, VersionThreadOccurrence[]>();
  private lengths = new Map<string, number>();
  private incoming: Record<LineageLevel, Map<string, LineageLink[]>> = { paragraph: new Map(), sentence: new Map() };
  private latestKnownVersion = -1;
  private latestThreadKeys = new Set<string>();
  private matches: Record<LineageLevel, WeakMap<LineageLink, VersionThreadOccurrence[]>> = { paragraph: new WeakMap(), sentence: new WeakMap() };

  constructor(versions: readonly SavedLineageVersion[], result: SavedVersionLineage) {
    versions.forEach((saved, version) => {
      if (saved.unavailable) return;
      if (!saved.commentsUnavailable) {
        this.latestKnownVersion = version;
        this.latestThreadKeys = new Set(saved.threads.map((thread) => thread.key));
      }
      for (const p of saved.paragraphs) this.lengths.set(`${version}:${p.index}`, p.text.length);
      for (const thread of saved.threads) {
        for (const range of thread.ranges) {
          const key = `${version}:${range.paragraph}`;
          const candidates = this.threadIndex.get(key) ?? [];
          candidates.push({ version, thread });
          this.threadIndex.set(key, candidates);
        }
      }
    });
    for (const comparison of result.comparisons) {
      if (comparison.unavailable || comparison.recordingGap) continue;
      for (const level of ["paragraph", "sentence"] as const) {
        const index = this.incoming[level];
        for (const link of comparison[level]) {
          if (link.kind !== "retained" && link.kind !== "rewritten" && link.kind !== "split" && link.kind !== "merged") continue;
          for (const target of link.to) {
            const key = `${target.version}:${target.paragraph}`;
            const incoming = index.get(key) ?? [];
            if (!incoming.includes(link)) incoming.push(link);
            index.set(key, incoming);
          }
        }
      }
    }
  }

  absentInLatest(occurrence: VersionThreadOccurrence): boolean {
    return this.latestKnownVersion > occurrence.version && !this.latestThreadKeys.has(occurrence.thread.key);
  }

  private at(range: LineageRange): VersionThreadOccurrence[] {
    const key = `${range.version}:${range.paragraph}`;
    const length = this.lengths.get(key) ?? 0;
    return (this.threadIndex.get(key) ?? []).filter((entry) => entry.thread.ranges.some((mark) => intersects(mark, range, length)));
  }

  private directRanges(link: LineageLink): LineageRange[] {
    return link.kind === "ambiguous" ? link.to : [...link.from, ...link.to];
  }

  /** Follow earlier passages without carrying a comment through an uncertain link or into a sibling split. */
  forRelationship(link: LineageLink, level: LineageLevel): VersionThreadOccurrence[] {
    const cached = this.matches[level].get(link);
    if (cached) return cached;
    const queue = this.directRanges(link);
    const checked = new Set<string>();
    const matches = new Map<SavedVersionThread, VersionThreadOccurrence>();
    for (let i = 0; i < queue.length; i++) {
      const range = queue[i];
      const key = rangeKey(range);
      if (checked.has(key)) continue;
      checked.add(key);
      for (const occurrence of this.at(range)) matches.set(occurrence.thread, occurrence);
      for (const prior of this.incoming[level].get(`${range.version}:${range.paragraph}`) ?? []) {
        const target = prior.to[0];
        if (prior.to.length === 1 && prior.from.length === 1 && target.paragraph === range.paragraph
          && range.start <= target.start && target.end <= range.end) {
          queue.push(prior.from[0]);
          continue;
        }
        for (const { from, to } of prior.unchanged) {
          if (to.paragraph !== range.paragraph) continue;
          const start = Math.max(range.start, to.start);
          const end = Math.min(range.end, to.end);
          if (start < end) queue.push({ ...from, start: from.start + (start - to.start), end: from.start + (end - to.start) });
        }
      }
    }
    const result = [...matches.values()].sort((a, b) => b.version - a.version);
    this.matches[level].set(link, result);
    return result;
  }
}
