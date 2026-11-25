import { commentReferences, omitCommentUrls } from "../../src/lib/commentReferences";
import {
  compareSavedVersions, lineageRangeText,
  type LineageLevel, type LineageLink, type LineageRange, type SavedParagraphRange, type SavedVersionLineage,
  type VersionComparison,
} from "../../src/lib/lineage";
import { LineageComments, type VersionThreadOccurrence } from "./lineageComments";
import {
  REPORT_LIMITS, reportWindow, type ReportChange, type ReportDraft, type ReportInterval,
  type ReportMessage, type ReportPair, type ReportPassage, type ReportText, type ReportThread, type SavedVersionReport,
} from "./reportData";
import type { AppLang, SavedLineageVersion, SavedVersionThread } from "./shared";

export interface SavedReportScope {
  versions: number[];
  initial: [number, number];
  excluded: ReadonlySet<string>;
  includeText: boolean;
  range: number;
  includeComments: boolean;
  includeDeleted: boolean;
  includeLinks: boolean;
  includeUnlocated: boolean;
  includeRelatedSources?: boolean;
  review?: boolean;
}

export interface SavedReportCandidate {
  id: string;
  kind: ReportChange["kind"];
  link: LineageLink;
  comments: VersionThreadOccurrence[];
  noteRange?: LineageRange;
}

interface StoredClip {
  start: number;
  end: number;
  number: number;
}

const options = { japaneseWordHints: true };

export function reportPairs(versions: readonly number[]): [number, number][] {
  return versions.flatMap((before, index) => versions.slice(index + 1).map((after): [number, number] => [before, after]));
}

function commonStart(names: string[]): string {
  const prefix = Array.from(names[0] ?? "");
  for (const name of names.slice(1)) {
    const points = Array.from(name);
    const end = prefix.findIndex((point, index) => point !== points[index]);
    if (end !== -1) prefix.length = end;
  }
  return prefix.join("");
}

export function savedReportTitle(filenames: string[]): string {
  const stems = filenames.map((name) => name.replace(/\.docx$/iu, "").normalize("NFC").trim() || name);
  const versions = stems.map((stem) => stem.replace(/[\s._-]+(?:(?:s|v|rev)\s*)?\d+(?:[._-]\d+)*$/iu, "") || stem);
  if (versions.every((name) => name === versions[0])) return versions[0] ?? "";
  const trim = (name: string) => name.replace(/^[\s._()-]+|[\s._()-]+$/gu, "");
  const prefix = trim(commonStart(versions));
  const reverse = (name: string) => Array.from(name).reverse().join("");
  const suffix = trim(reverse(commonStart(versions.map(reverse))));
  for (const part of [prefix, suffix]) {
    if (Array.from(part).length > 1 && !/^\d[\d\s.-]*$/u.test(part)) return part;
  }
  return [...new Set(stems)].join(" / ");
}

const threadContent = (thread: SavedVersionThread) => JSON.stringify(thread.detail.comments.map((c) =>
  [c.key, c.author, c.dateUtc ?? c.date, c.done, c.text, c.links]));

function shiftRange(range: LineageRange, before: number, after: number): LineageRange {
  return { ...range, version: range.version === 0 ? before : after };
}

class SourceText {
  text = "";
  private paragraphs = new Map<number, { start: number; length: number; heading?: string }>();
  private pages: number[] = [];
  private pageMap?: { estimated: boolean };

  constructor(version: SavedLineageVersion) {
    if (version.unavailable) return;
    let heading: string | undefined;
    const text: string[] = [];
    let offset = 0;
    for (const p of version.paragraphs) {
      if (p.kind === "heading") heading = Array.from(p.text).slice(0, 80).join("");
      this.paragraphs.set(p.index, { start: offset, length: p.text.length, heading });
      text.push(p.text);
      offset += p.text.length + 2;
    }
    this.text = text.join("\n\n");
    if (version.pageMap) {
      this.pageMap = version.pageMap;
      this.pages = version.pageMap.breaks.flatMap((b) => {
        const p = this.paragraphs.get(b.paragraph);
        return p ? [p.start + b.offset] : [];
      });
    }
  }

  interval(range: SavedParagraphRange): ReportInterval | undefined {
    const p = this.paragraphs.get(range.paragraph);
    if (p && range.start >= 0 && range.end >= range.start && range.end <= p.length) {
      return [p.start + range.start, p.start + range.end];
    }
  }

  heading(paragraph: number): string | undefined {
    return this.paragraphs.get(paragraph)?.heading;
  }

  context(focus: number, limit: number): { bounds: ReportInterval; text: ReportText } {
    const next = this.pageMap && this.pages.findIndex((offset) => offset > focus);
    const page = next === undefined ? undefined : next === -1 ? this.pages.length : next;
    const left = page ? this.pages[page - 1] : 0;
    // A paragraph separator belongs to neither of the two pages.
    const boundary = page === undefined ? undefined : this.pages[page];
    const right = boundary === undefined ? this.text.length :
      boundary > 1 && this.text.slice(boundary - 2, boundary) === "\n\n" ? boundary - 2 : boundary;
    const [start, end] = reportWindow(this.text, focus, limit === Infinity && page === undefined ? 400 : limit, left, right);
    return {
      bounds: [start, end],
      text: {
        text: this.text.slice(start, end),
        beforeCut: start > left,
        afterCut: end < right,
        page: page === undefined ? undefined : { number: page + 1, estimated: this.pageMap!.estimated },
      },
    };
  }
}

/** Index the observed source once and compute nonadjacent comparisons on demand. */
export class SavedReportAnalysis {
  private adjacentComments: LineageComments;
  private candidates = new Map<string, { pair: VersionComparison; changes: SavedReportCandidate[] }>();
  private sources = new Map<number, SourceText>();
  private counts = new Map<AppLang, Array<{ characters: number; words: number } | undefined>>();

  constructor(
    private versions: readonly SavedLineageVersion[],
    private lineage: SavedVersionLineage,
    readonly level: LineageLevel,
  ) {
    this.adjacentComments = new LineageComments(versions, lineage);
  }

  private source(index: number): SourceText {
    let source = this.sources.get(index);
    if (!source) {
      source = new SourceText(this.versions[index]);
      this.sources.set(index, source);
    }
    return source;
  }

  private compare(before: number, after: number): VersionComparison {
    if (after === before + 1) return this.lineage.comparisons[before];
    const pair = compareSavedVersions([this.versions[before], this.versions[after]], options).comparisons[0];
    const shift = (r: LineageRange) => shiftRange(r, before, after);
    const links = pair[this.level].map((link): LineageLink => ({
      ...link, from: link.from.map(shift), to: link.to.map(shift),
      unchanged: link.unchanged.map((match) => ({ from: shift(match.from), to: shift(match.to) })),
    }));
    return {
      ...pair, before, after, [this.level]: links,
      recordingGap: this.versions.slice(before + 1, after + 1).some((v) => v.gapBefore || v.unavailable),
    };
  }

  changes(before: number, after: number): readonly SavedReportCandidate[] {
    return this.analyze(before, after).changes;
  }

  describe(candidate: SavedReportCandidate): string {
    const range = candidate.link.to[0] ?? candidate.link.from[0];
    if (!range) return "";
    const sample = Array.from(lineageRangeText(this.versions, range));
    const heading = this.source(range.version).heading(range.paragraph);
    return [heading, sample.slice(0, 75).join("") + (sample.length > 75 ? "…" : "")].filter(Boolean).join(" · ");
  }

  private draft(index: number, lang: AppLang, labels: readonly string[], unlocated: number[] = []): ReportDraft {
    if (!this.counts.has(lang)) {
      const segmenter = new Intl.Segmenter(lang, { granularity: "word" });
      this.counts.set(lang, this.versions.map((version) => {
        if (version.unavailable) return;
        let characters = 0;
        let words = 0;
        for (const paragraph of version.paragraphs) {
          characters += Array.from(paragraph.text).length;
          for (const segment of segmenter.segment(paragraph.text)) if (segment.isWordLike) words++;
        }
        return { characters, words };
      }));
    }
    const version = this.versions[index];
    return {
      label: labels[index], filename: version.label, createdAt: version.fileCreatedAt, archivedAt: version.archivedAt,
      modifiedAt: version.sourceModified, size: version.fileSize, ...this.counts.get(lang)?.[index],
      unavailable: Boolean(version.unavailable), unlocated,
    };
  }

  drafts(lang: AppLang, labels: readonly string[]): ReportDraft[] {
    return this.versions.map((_, index) => this.draft(index, lang, labels));
  }

  private analyze(before: number, after: number) {
    const key = `${before}:${after}`;
    const stored = this.candidates.get(key);
    if (stored) return stored;
    const pair = this.compare(before, after);
    const changes = pair[this.level].flatMap((link, index): SavedReportCandidate[] => {
      let kind: SavedReportCandidate["kind"] = link.kind;
      let noteRange: LineageRange | undefined;
      if (link.kind === "retained" && !link.reordered) {
        if (this.versions[before].commentsUnavailable || this.versions[after].commentsUnavailable) return [];
        const old = this.adjacentComments.atRanges(link.from);
        const current = this.adjacentComments.atRanges(link.to);
        const oldContent = new Set(old.map((entry) => threadContent(entry.thread)));
        const currentContent = new Set(current.map((entry) => threadContent(entry.thread)));
        const changed = current.find((entry) => !oldContent.has(threadContent(entry.thread))) ??
          old.find((entry) => !currentContent.has(threadContent(entry.thread)));
        if (!changed) return [];
        const range = changed.thread.ranges.find((mark) => [...link.from, ...link.to].some((target) =>
          target.version === changed.version && target.paragraph === mark.paragraph && mark.start <= target.end && mark.end >= target.start));
        if (range) noteRange = { ...range, version: changed.version };
        kind = "comment";
      }
      return [{ id: `${key}:${index}`, kind, link, comments: this.adjacentComments.forRelationship(link, this.level), noteRange }];
    });
    const analyzed = { pair, changes };
    this.candidates.set(key, analyzed);
    return analyzed;
  }

  project(scope: SavedReportScope, lang: AppLang, labels: readonly string[]): SavedVersionReport {
    const selected = [...new Set(scope.versions)].sort((a, b) => a - b);
    const ids = new Map(selected.map((version, index) => [version, index]));
    const beforeInitial = ids.get(scope.initial[0]);
    const afterInitial = ids.get(scope.initial[1]);
    const initial = beforeInitial === undefined || afterInitial === undefined
      ? selected.slice(0, 2).map((_, index) => index) : [beforeInitial, afterInitial];
    const clips: ReportText[] = [];
    const stored = new Map<number, Map<string, StoredClip>>();
    const threads: ReportThread[] = [];
    const threadIds = new Map<SavedVersionThread, number>();
    const pairThreads: { change: ReportChange; occurrences: VersionThreadOccurrence[] }[] = [];
    const latest = [...selected].reverse().find((index) => !this.versions[index].unavailable && !this.versions[index].commentsUnavailable);
    const latestKeys = new Set(latest === undefined ? [] : this.versions[latest].threads.map((thread) => thread.key));
    const includeText = scope.includeText;

    const passage = (range: LineageRange, candidate: SavedReportCandidate, side: "from" | "to"): ReportPassage => {
      const view: ReportPassage = { paragraph: range.paragraph + 1 };
      if (!includeText) return view;
      const source = this.source(range.version);
      const interval = source.interval(range);
      if (!interval) return view;
      view.heading = source.heading(range.paragraph);
      const [start, end] = interval;
      const { link, noteRange } = candidate;
      const identical = link.unchanged.map((span) => span[side])
        .filter((r) => r.paragraph === range.paragraph && r.version === range.version && r.start >= range.start && r.end <= range.end)
        .map((span) => source.interval(span)!)
        .sort((a, b) => a[0] - b[0]);
      const edits: ReportInterval[] = [];
      if (link.kind !== "ambiguous" && link.kind !== "retained" && !link.alignmentPartial) {
        let offset = start;
        for (const [left, right] of identical) {
          if (offset < left) edits.push([offset, left]);
          offset = Math.max(offset, right);
        }
        if (offset < end) edits.push([offset, end]);
      }
      const note = noteRange?.version === range.version && noteRange.paragraph === range.paragraph
        ? source.interval(noteRange)?.[0] : undefined;
      const marks = link.alignmentPartial ? identical : edits;
      const focus = marks[0]?.[0] ?? Math.max(start, Math.min(end, note ?? start + Math.min(10, Math.floor((end - start) / 2))));
      const context = source.context(focus, REPORT_LIMITS[scope.range]);
      const [left, right] = context.bounds;
      let byWindow = stored.get(range.version);
      if (!byWindow) stored.set(range.version, byWindow = new Map());
      const key = `${left}:${right}`;
      let clip = byWindow.get(key);
      if (!clip) {
        clip = { start: left, end: right, number: clips.push(context.text) - 1 };
        byWindow.set(key, clip);
      }
      view.excerpt = {
        clip: clip.number, focus: Math.max(0, Math.min(right - left, focus - left)),
        marks: marks.filter(([a, b]) => a < right && b > left).map(([a, b]) => [Math.max(left, a) - left, Math.min(right, b) - left]),
      };
      return view;
    };

    const pairs: ReportPair[] = reportPairs(selected).map(([before, after]) => {
      const { pair, changes } = this.analyze(before, after);
      const included = changes.filter((candidate) => !scope.excluded.has(candidate.id)).map((candidate): ReportChange => {
        const { link } = candidate;
        const within = new Set([before, after]);
        const omitPast = link.kind === "reappeared" && link.from.some((r) => !ids.has(r.version));
        const item: ReportChange = {
          ...(scope.review ? { reviewId: candidate.id } : {}),
          kind: omitPast ? "added" : candidate.kind,
          reordered: Boolean(link.reordered), alignmentPartial: Boolean(link.alignmentPartial),
          evidence: link.evidence.filter((e) => !omitPast || e !== "past"),
          candidateCount: link.candidateCount,
          before: link.from.filter((r) => within.has(r.version)).map((r) => passage(r, candidate, "from")),
          after: link.to.filter((r) => r.version === after).map((r) => passage(r, candidate, "to")),
          threads: [],
        };
        if (scope.includeComments) pairThreads.push({ change: item, occurrences: candidate.comments });
        return item;
      });
      return { before: ids.get(before)!, after: ids.get(after)!, gap: pair.recordingGap, unavailable: pair.unavailable, changes: included };
    });

    const canShareThread = ({ version, thread }: VersionThreadOccurrence) =>
      (ids.has(version) || scope.includeRelatedSources) && (scope.includeDeleted || !this.adjacentComments.absentInLatest({ version, thread }));

    const threadId = ({ version, thread }: VersionThreadOccurrence): number => {
      const existing = threadIds.get(thread);
      if (existing !== undefined) return existing;
      const originals = thread.ranges.map((range): ReportPassage | undefined => {
        const source = this.source(version);
        const interval = source.interval(range);
        if (!interval) return;
        const [start, end] = interval;
        let clip = [...(stored.get(version)?.values() ?? [])].find((part) => part.start <= start && part.end >= end);
        if (!clip && includeText && scope.includeRelatedSources) {
          const focus = start + Math.floor((end - start) / 2);
          const context = source.context(focus, REPORT_LIMITS[scope.range]);
          const [left, right] = context.bounds;
          if (left <= start && right >= end) clip = { start: left, end: right, number: clips.push(context.text) - 1 };
        }
        if (!clip) return;
        return {
          paragraph: range.paragraph + 1, heading: source.heading(range.paragraph),
          excerpt: { clip: clip.number, focus: start - clip.start, marks: [[start - clip.start, end - clip.start]], point: start === end },
        };
      });
      const messages = thread.detail.comments.map((comment): ReportMessage => {
        const { tags, links } = commentReferences(comment.text, comment.links);
        const share = (value: string) => scope.includeLinks ? value : omitCommentUrls(value);
        const date = comment.dateUtc ?? comment.date;
        return { author: comment.author && share(comment.author), date: date && Number.isFinite(Date.parse(date)) ? date : undefined,
          text: share(comment.text), tags, links: scope.includeLinks ? links : undefined };
      });
      const draft = ids.get(version);
      const id = threads.push({
        draft, sourceLabel: draft === undefined ? labels[version] : undefined,
        ...(scope.review ? { reviewVersion: version } : {}), resolved: Boolean(thread.detail.comments[0]?.done),
        absent: scope.includeRelatedSources ? this.adjacentComments.absentInLatest({ version, thread })
          : latest !== undefined && latest > version && !latestKeys.has(thread.key),
        unavailable: thread.detail.versionUnavailable, messages,
        original: originals.length && originals.every((range) => range !== undefined) ? originals : undefined,
      }) - 1;
      threadIds.set(thread, id);
      return id;
    };

    for (const { change, occurrences } of pairThreads) change.threads = occurrences.filter(canShareThread).map(threadId);
    const drafts: ReportDraft[] = selected.map((version) => {
      const item = this.versions[version];
      const unlocated = scope.includeComments && scope.includeUnlocated ? item.threads.filter((thread) => !thread.ranges.length)
        .map((thread) => ({ version, thread })).filter(canShareThread).map(threadId) : [];
      return this.draft(version, lang, labels, unlocated);
    });
    const title = savedReportTitle(selected.map((index) => {
      const version = this.versions[index];
      // Watched copies include an archive hash and sequence number after the document's filename.
      return version.source === "archive" ? version.label.replace(/-[0-9a-f]{10}-s\d+\.docx$/iu, "") : version.label;
    }));
    return { title, lang, level: this.level, initial, range: scope.range, includeText, drafts, clips, threads, pairs };
  }
}
