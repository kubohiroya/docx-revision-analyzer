import { commentReferences } from "../../src/lib/commentReferences";
import type { SavedParagraphRange } from "../../src/lib/lineage";
import type { RevisionOccurrence } from "../../src/lib/revisionHistory";
import type { SavedLineageVersion, SavedVersionThread } from "./shared";

export interface HistoryEntry {
  id: string;
  foldedText: string;
  foldedQuote: string;
  version: number;
  kind: RevisionOccurrence["kind"] | "comment" | "reply";
  text: string;
  quote?: string;
  author?: string;
  date?: string;
  tags: string[];
  ranges: SavedParagraphRange[];
  thread?: SavedVersionThread;
  revision?: RevisionOccurrence;
  resolved?: boolean;
  absentAt?: number;
  threadGroup?: string;
}

export function historyEntries(version: SavedLineageVersion, index: number, revisions: readonly RevisionOccurrence[]): HistoryEntry[] {
  const history: HistoryEntry[] = revisions.map(revision => ({
    id: `${index}:revision:${revision.occurrence}`,
    foldedText: searchKey(revision.text), foldedQuote: "",
    version: index, kind: revision.kind, text: revision.text, author: revision.author, date: revision.date,
    ranges: revision.ranges, tags: [], revision,
  }));
  for (const [threadIndex, thread] of version.threads.entries()) {
    const quote = thread.detail.passage?.quote ?? thread.detail.recordedAnchor?.quote;
    const foldedQuote = searchKey(quote ?? "");
    thread.detail.comments.forEach((comment, i) => {
      history.push({
        id: `${index}:thread:${threadIndex}:${i}`, foldedText: searchKey(comment.text), foldedQuote,
        version: index, kind: i ? "reply" : "comment", text: comment.text, author: comment.author,
        date: comment.dateUtc ?? comment.date, tags: commentReferences(comment.text, comment.links).tags,
        ranges: thread.ranges, thread, resolved: thread.detail.comments[0].done,
        quote,
      });
    });
  }
  return history;
}

interface HistoryVersion {
  version: SavedLineageVersion;
  entries: HistoryEntry[];
  changesRead: boolean;
  commentsRead: boolean;
}

function threadKeys(threads: readonly SavedVersionThread[]): Map<string, SavedVersionThread | null> {
  const keys = new Map<string, SavedVersionThread | null>();
  for (const thread of threads) {
    if (thread.key.startsWith("d:")) keys.set(thread.key, keys.has(thread.key) ? null : thread);
  }
  return keys;
}

/** Reconcile observations only between adjacent, readable versions of the selected path. */
export function reconcileHistoryVersions(earlier: HistoryVersion, later: HistoryVersion, index: number): void {
  const before = earlier.version, after = later.version;
  if (before.unavailable || after.unavailable || after.gapBefore) return;
  if (earlier.commentsRead && later.commentsRead) {
    const previous = threadKeys(before.threads);
    const next = threadKeys(after.threads);
    const groupIds = new Map<string, string>();
    for (const entry of earlier.entries) {
      const original = entry.thread;
      if (!original || previous.get(original.key) !== original) continue;
      const key = original.key;
      const laterThread = next.get(key);
      if (laterThread === undefined) entry.absentAt = index;
      if (!laterThread) continue;
      const first = original.detail.comments[0], current = laterThread.detail.comments[0];
      const firstDate = first.dateUtc ?? first.date, currentDate = current.dateUtc ?? current.date;
      if (firstDate && currentDate && firstDate !== currentDate) continue;
      if (first.author && current.author && first.author !== current.author) continue;
      entry.threadGroup ??= groupIds.get(key) ?? entry.id;
      groupIds.set(key, entry.threadGroup);
    }
    for (const entry of later.entries) {
      if (entry.thread) entry.threadGroup = groupIds.get(entry.thread.key);
    }
  }
  if (!earlier.changesRead || !later.changesRead) return;
  const tracked = (entry: HistoryEntry) => entry.revision && [entry.kind, entry.revision.wordId,
    entry.date, entry.author, entry.text].join("\0");
  const counts = new Map<string, number>();
  for (const entry of earlier.entries) {
    const key = tracked(entry);
    if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  const currentIds = new Set(later.entries.flatMap(entry => entry.revision?.wordId ? [entry.revision.wordId] : []));
  const currentText = new Set(later.entries.flatMap(entry => entry.revision ? [entry.text] : []));
  const originalParagraphs = new Map(before.paragraphs.map(paragraph => [paragraph.index, paragraph]));
  const currentParagraphs = new Map<string, number>();
  for (const paragraph of after.paragraphs) {
    if (!paragraph.paraId) continue;
    const key = `${paragraph.paraId}\0${paragraph.text}`;
    currentParagraphs.set(key, (currentParagraphs.get(key) ?? 0) + 1);
  }
  for (const entry of earlier.entries) {
    const revision = entry.revision;
    const key = tracked(entry);
    if (!key || !revision?.wordId || !entry.ranges.length || counts.get(key) !== 1 ||
        currentIds.has(revision.wordId) || currentText.has(entry.text)) continue;
    const aligned = entry.ranges.every(range => {
      const source = originalParagraphs.get(range.paragraph);
      return source?.paraId && currentParagraphs.get(`${source.paraId}\0${source.text}`) === 1;
    });
    if (aligned) entry.absentAt = index;
  }
}

export function groupedHistoryRows(entries: readonly HistoryEntry[]): HistoryEntry[][] {
  const rows: HistoryEntry[][] = [];
  const grouped = new Map<string, HistoryEntry[]>();
  for (const entry of entries) {
    const group = entry.threadGroup;
    const row = group && grouped.get(group);
    if (row) row.push(entry);
    else {
      const peers = [entry];
      if (group) grouped.set(group, peers);
      rows.push(peers);
    }
  }
  return rows.flatMap(row => row.every(entry => entry.version === row[0].version) ? row.map(entry => [entry]) : [row]);
}

export function searchKey(text: string): string {
  return text.normalize("NFC").replace(/[A-Z]/g, letter => letter.toLowerCase());
}

const graphemes = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** Match normalized text while keeping highlights on complete original graphemes. */
export function historyMatch(text: string, query: string): { start: number; end: number } | undefined {
  if (!query) return undefined;
  const normalized = searchKey(text);
  const offset = normalized.indexOf(query);
  if (offset < 0) return undefined;
  if (normalized.length === text.length && normalized === text.replace(/[A-Z]/g, letter => letter.toLowerCase())) {
    const first = graphemes.segment(text).containing(offset);
    const last = graphemes.segment(text).containing(offset + query.length - 1);
    if (first && last) return { start: first.index, end: last.index + last.segment.length };
  }
  let cursor = 0;
  let start: number | undefined;
  for (const { segment, index } of graphemes.segment(text)) {
    cursor += searchKey(segment).length;
    if (cursor > offset && start === undefined) start = index;
    if (cursor >= offset + query.length) return { start: start ?? index, end: index + segment.length };
  }
  return undefined;
}
