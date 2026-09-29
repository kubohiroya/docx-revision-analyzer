/**
 * insertionKinds.ts
 *
 * 挿入を「一括挿入」「細かい編集」「移動・並べ替え」に分類する共通の規則。
 * docx-revision-chart (積み上げ棒) と docx-revision-flow (段落の色) で同じ規則を使う。
 *
 *  - 移動・並べ替え (moved): 挿入された文章 (RELOCATION_MIN_CHARS 文字以上) と同じ内容が、
 *    文書のどこかで削除されている。コピー＋貼り付け＋削除や、Word が移動として記録しなかった
 *    カット＋貼り付けがこれにあたる (Word が移動として記録したもの w:moveFrom / w:moveTo は、
 *    そもそも挿入・削除のイベントには含まれない)。
 *  - 一括挿入 (bulk): 同じ作成者・同じ日時の挿入 (移動・並べ替えを除く) の合計が bulkChars 文字以上。
 *    Word は貼り付けた文章を段落ごとに別々の w:ins として記録するが、日時は同じになるため。
 *  - 細かい編集 (fine): それ以外。
 */

import type { RevisionEvent } from "./docxRevisions";

export type InsertionKind = "bulk" | "fine" | "moved";

/** 同じ作成者・同じ日時の挿入の合計がこれ以上なら一括挿入とみなす (文字数) の既定値 */
export const DEFAULT_BULK_CHARS = 150;

/** これ未満の文字数の挿入・削除は、並べ替えの判定に使わない (偶然の一致を避ける) */
export const RELOCATION_MIN_CHARS = 20;

/** 並べ替えの判定で、空白の違いを無視するための正規化 */
export const normalizeForMatch = (t: string) => t.replace(/\s+/g, "");

/** 文章 text が、corpus (正規化済みの文章を連結したもの) に含まれているか */
export function matchesCorpus(text: string, corpus: string): boolean {
  const n = normalizeForMatch(text);
  return n.length >= RELOCATION_MIN_CHARS && corpus.includes(n);
}

/**
 * 挿入イベントの insKind を設定する (削除イベントはそのまま)。
 * 並べ替えの判定には、全期間の削除イベントの文章を使う。
 */
export function classifyInsertions(events: RevisionEvent[], bulkChars = DEFAULT_BULK_CHARS): void {
  const delCorpus = events
    .filter((e) => e.type === "del")
    .map((e) => normalizeForMatch(e.text))
    .join("");
  const groupKey = (e: RevisionEvent) => `${e.author ?? ""}|${e.date.getTime()}`;
  const groups = new Map<string, number>();
  for (const e of events) {
    if (e.type !== "ins") continue;
    if (matchesCorpus(e.text, delCorpus)) {
      e.insKind = "moved";
    } else {
      groups.set(groupKey(e), (groups.get(groupKey(e)) ?? 0) + e.chars);
    }
  }
  for (const e of events) {
    if (e.type !== "ins" || e.insKind === "moved") continue;
    e.insKind = (groups.get(groupKey(e)) ?? 0) >= bulkChars ? "bulk" : "fine";
  }
}
