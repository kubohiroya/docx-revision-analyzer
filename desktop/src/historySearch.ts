import { commentReferences } from "../../src/lib/commentReferences";
import type { SavedParagraphRange } from "../../src/lib/lineage";
import type { RevisionOccurrence } from "../../src/lib/revisionHistory";
import type { SavedLineageVersion, SavedVersionThread } from "./shared";

export interface HistoryEntry {
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
}

export function historyEntries(version: SavedLineageVersion, index: number, revisions: readonly RevisionOccurrence[]): HistoryEntry[] {
  const history: HistoryEntry[] = revisions.map(revision => ({
    version: index, kind: revision.kind, text: revision.text, author: revision.author, date: revision.date,
    ranges: revision.ranges, tags: [], revision,
  }));
  for (const thread of version.threads) {
    thread.detail.comments.forEach((comment, i) => {
      history.push({
        version: index, kind: i ? "reply" : "comment", text: comment.text, author: comment.author,
        date: comment.dateUtc ?? comment.date, tags: commentReferences(comment.text, comment.links).tags,
        ranges: thread.ranges, thread, resolved: thread.detail.comments[0].done,
        quote: thread.detail.passage?.quote ?? thread.detail.recordedAnchor?.quote,
      });
    });
  }
  return history;
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
