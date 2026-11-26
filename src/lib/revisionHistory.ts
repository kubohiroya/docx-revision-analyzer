import { parseDocxLayout, type LayoutRevisionKind, type LayoutRevisionMark, type RevRef, type RevisionExcludedRegion } from "./docxLayout";
import type { DocxInput } from "./input";
import { savedSegmentText, type SavedParagraphRange } from "./lineage";

export interface RevisionOccurrence {
  part: "word/document.xml";
  occurrence: number;
  wordId?: string;
  kind: LayoutRevisionKind;
  text: string;
  author?: string;
  date?: string;
  moveName?: string;
  moveOccurrence?: number;
  paragraphMark?: boolean;
  ranges: SavedParagraphRange[];
  context?: { prefix: string; suffix: string };
}

export interface RevisionIndex {
  records: RevisionOccurrence[];
  excluded: RevisionExcludedRegion[];
}

export type RevisionHistoryRead =
  | ({ status: "read" } & RevisionIndex)
  | { status: "changed" | "unavailable" | "failed" | "disabled"; records: [] };

/** Extract individual XML occurrences, including nested edits and repeated Word IDs. */
export async function readRevisionIndex(input: DocxInput): Promise<RevisionIndex> {
  const marks = new Map<number, LayoutRevisionMark[]>();
  const excluded = new Set<RevisionExcludedRegion>();
  const model = await parseDocxLayout(input, undefined, undefined, { marks, excluded });
  const records = new Map<RevRef, RevisionOccurrence>();
  for (const p of model.paras) {
    const selected = marks.get(p.index);
    if (p.kind === "table" || !selected?.length) continue;
    let visible = "";
    const offsets = [0];
    for (const segment of p.segs) {
      if (!segment.del) visible += savedSegmentText(segment);
      offsets.push(visible.length);
    }
    for (const mark of selected) {
      const { ref, kind, start, end, block } = mark;
      let record = records.get(ref);
      const continues = record !== undefined;
      if (!record) {
        record = {
          part: "word/document.xml", occurrence: records.size,
          wordId: ref.id || undefined, kind, text: "", author: ref.author,
          date: ref.date && Number.isFinite(ref.date.getTime()) ? ref.date.toISOString() : undefined,
          moveName: ref.moveName, paragraphMark: mark.paragraphMark, ranges: [],
        };
        records.set(ref, record);
      }
      if (continues) record.text += "\n";
      record.text += p.segs.slice(start, end).map(savedSegmentText).join("");
      if (!visible.trim() || (block && (kind === "del" || kind === "moveFrom"))) continue;
      const from = offsets[start], to = offsets[end];
      if (!record.context) record.context = { prefix: visible.slice(Math.max(0, from - 40), from), suffix: visible.slice(to, to + 40) };
      if (kind === "del" || kind === "moveFrom" || from === to) {
        record.ranges.push({ paragraph: p.index, start: from, end: from });
        continue;
      }
      let current: SavedParagraphRange | undefined;
      for (let index = start; index < end; index++) {
        if (p.segs[index].del) {
          current = undefined;
        } else if (offsets[index] < offsets[index + 1]) {
          if (current) current.end = offsets[index + 1];
          else {
            current = { paragraph: p.index, start: offsets[index], end: offsets[index + 1] };
            record.ranges.push(current);
          }
        }
      }
    }
  }
  const moves = new Map<string, RevisionOccurrence[]>();
  for (const record of records.values()) {
    if (!record.moveName) continue;
    const group = moves.get(record.moveName) ?? [];
    group.push(record);
    moves.set(record.moveName, group);
  }
  for (const group of moves.values()) {
    if (group.length !== 2 || group[0].kind === group[1].kind) continue;
    group[0].moveOccurrence = group[1].occurrence;
    group[1].moveOccurrence = group[0].occurrence;
  }
  return { records: [...records.values()], excluded: [...excluded] };
}
