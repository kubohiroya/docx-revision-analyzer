import type { CommentLink } from "../../src/lib/commentReferences";
import type { LineageKind, LineageLevel, MatchEvidence } from "../../src/lib/lineage";
import type { AppLang } from "./shared";

export const REPORT_LIMITS = [20, 80, 200, 400, Infinity] as const;
export const DEFAULT_REPORT_RANGE = 2;

export interface ReportText {
  text: string;
  beforeCut: boolean;
  afterCut: boolean;
  page?: { number: number; estimated: boolean };
}

export type ReportInterval = [start: number, end: number];

export interface ReportPassage {
  paragraph: number;
  heading?: string;
  excerpt?: { clip: number; focus: number; marks: ReportInterval[]; point?: boolean };
}

export interface ReportMessage {
  author?: string;
  date?: string;
  text: string;
  tags: string[];
  links?: CommentLink[];
}

export interface ReportThread {
  draft: number;
  resolved: boolean;
  absent: boolean;
  unavailable: boolean;
  messages: ReportMessage[];
  original?: ReportPassage[];
}

export interface ReportChange {
  kind: LineageKind | "comment";
  reordered: boolean;
  evidence: MatchEvidence[];
  alignmentPartial: boolean;
  candidateCount?: number;
  before: ReportPassage[];
  after: ReportPassage[];
  threads: number[];
}

export interface ReportDraft {
  label: string;
  archivedAt?: string;
  modifiedAt?: string;
  unavailable: boolean;
  unlocated: number[];
}

export interface ReportPair {
  before: number;
  after: number;
  gap: boolean;
  unavailable: boolean;
  changes: ReportChange[];
}

export interface SavedVersionReport {
  lang: AppLang;
  level: LineageLevel;
  initial: [number, number];
  range: number;
  includeText: boolean;
  drafts: ReportDraft[];
  clips: ReportText[];
  threads: ReportThread[];
  pairs: ReportPair[];
}

// The center stays fixed as the visible context grows. Keep UTF-16 boundaries away from split surrogate pairs.
export function reportWindow(text: string, focus: number, limit: number, left = 0, right = text.length): ReportInterval {
  if (right - left <= limit) return [left, right];
  let start = Math.max(left, Math.min(focus - Math.floor(limit / 2), right - limit));
  let end = start + limit;
  const lowSurrogate = (index: number) => text.charCodeAt(index) >= 0xdc00 && text.charCodeAt(index) <= 0xdfff;
  if (lowSurrogate(start)) start++;
  if (lowSurrogate(end)) end--;
  return [start, end];
}
