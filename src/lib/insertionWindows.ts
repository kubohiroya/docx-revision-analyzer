/**
 * insertionWindows.ts
 *
 * 位置付きの編集イベント列 (revisionPositions.ts) から、「短い時間に、文書の狭い範囲へ
 * まとまって行われた挿入」を窓 (まとまり) として取り出し、窓ごとに特徴量を計算する。
 * 特徴量は判定 (しきい値) から独立しており、判定ルール (insertionRules.ts) から参照する。
 *
 * まとめ方:
 *  - 対象は日時のある挿入。移動 (w:moveTo) と並べ替え (同じ内容が文書のどこかで削除されている
 *    挿入。insertionKinds.ts と同じ規則) は除く。
 *  - 挿入を時刻順に見ていき、開始から seconds 秒以内の窓のうち、最終文書での距離が
 *    chars 文字以内 (paras を指定した場合は、段落インデックスの差も paras 以内) のものに加える。
 *    当てはまる窓が無ければ新しい窓を開く。窓の時間幅は seconds を超えない
 *    (細かい入力が続いても、長い時間の挿入が1つの窓に連なることはない)。
 *  - 位置は最終文書 (すべての変更を反映した状態) の座標で測る。後で削除された挿入は、
 *    文字が消えた地点にあるものとして扱う。
 *
 * 時刻の解像度:
 *  - Word の w:date は分単位 (秒が常に 0) で記録されることが多い。文書中のすべての日時から
 *    解像度 (60秒 / 1秒 / 1ミリ秒) を推定し、durationSec には解像度ぶんの幅 (1秒未満なら1秒) を
 *    足す。同じ記録時刻の挿入でも、実際には解像度の幅のどこかで行われたとみなすため。
 */

import type { PositionedRevisionEvent, RevisionRange } from "./revisionPositions";
import { matchesCorpus, normalizeForMatch } from "./insertionKinds";

export interface WindowOptions {
  /** 時間窓 Δt (秒)。0 なら同じ記録時刻の挿入だけをまとめる */
  seconds: number;
  /** 文書上の距離 L (文字)。省略すると距離を問わない */
  chars?: number;
  /** 文書上の距離 P (段落)。指定した場合は chars と両方を満たす必要がある */
  paras?: number;
  /** true なら同じ作成者の挿入だけをまとめる */
  byAuthor?: boolean;
}

/**
 * durationSec に足す幅の下限 (秒)。Word は入力中の文字列にも1つの日時しか付けないため、
 * 時刻がミリ秒まで記録されていても、1件の挿入にかかった時間は分からない
 */
export const MIN_DURATION_SEC = 1;

export const DEFAULT_WINDOW_OPTIONS: WindowOptions = { seconds: 60, chars: 2000 };

export interface WindowFeatures {
  /** 挿入文字数の合計 */
  insertedChars: number;
  /** 時間幅 (秒)。記録時刻の幅に、時刻の解像度 (1秒未満なら1秒) を足したもの */
  durationSec: number;
  /** 挿入速度 (文字/秒) = insertedChars / durationSec */
  cps: number;
  /** 最終文書での範囲の文字数 */
  spanChars: number;
  /** 最終文書での範囲の段落数 (最初と最後の段落の間をすべて含む) */
  spanParas: number;
  /** 最大の単一の挿入 (w:ins) の文字数 */
  maxSingleInsert: number;
  /** 挿入が及ぶ段落の数 */
  paraCount: number;
  /** 窓の後に、その範囲へ加えられた挿入・削除の文字数 ÷ insertedChars */
  postEditRatio: number;
  /** 窓の直前 (seconds 以内) から窓の終わりまでに、同じ範囲で削除があったか (置き換え) */
  precededByDeletion: boolean;
  /** その削除の文字数 */
  precedingDeletedChars: number;
  /** 挿入の件数 */
  insertCount: number;
}

export type FeatureName = keyof WindowFeatures;

export interface InsertionWindow {
  /** 0始まりの通し番号 (開始時刻の順) */
  index: number;
  start: Date;
  end: Date;
  authors: string[];
  /** 窓に含まれる挿入の w:id */
  eventIds: string[];
  /** 最終文書での範囲 (start は最も前、end は最も後ろの位置) */
  range: RevisionRange;
  features: WindowFeatures;
}

export interface InsertionWindowResult {
  options: WindowOptions;
  /** 推定した時刻の解像度 (秒) */
  timeResolutionSec: number;
  windows: InsertionWindow[];
}

/**
 * 日時の列から時刻の解像度 (秒) を推定する。
 * すべて秒が 0 なら 60、ミリ秒がすべて 0 なら 1、それ以外は 0.001。
 */
export function estimateTimeResolutionSec(dates: Date[]): number {
  if (dates.length === 0) return 1;
  if (dates.every((d) => d.getTime() % 60000 === 0)) return 60;
  if (dates.every((d) => d.getTime() % 1000 === 0)) return 1;
  return 0.001;
}

/** 窓の対象とする挿入 (日時あり・移動でも並べ替えでもない) */
export function windowCandidates(events: PositionedRevisionEvent[]): PositionedRevisionEvent[] {
  const delCorpus = events
    .filter((e) => e.type === "del" && !e.move)
    .map((e) => normalizeForMatch(e.text))
    .join("");
  return events.filter(
    (e) => e.type === "ins" && !e.move && e.date && e.chars > 0 && !matchesCorpus(e.text, delCorpus)
  );
}

/** 2つの範囲の、最終文書での距離 (重なれば 0) */
function rangeGap(a: RevisionRange, b: RevisionRange): number {
  if (a.end.docOffset < b.start.docOffset) return b.start.docOffset - a.end.docOffset;
  if (b.end.docOffset < a.start.docOffset) return a.start.docOffset - b.end.docOffset;
  return 0;
}

function paraGap(a: RevisionRange, b: RevisionRange): number {
  if (a.end.paraIndex < b.start.paraIndex) return b.start.paraIndex - a.end.paraIndex;
  if (b.end.paraIndex < a.start.paraIndex) return a.start.paraIndex - b.end.paraIndex;
  return 0;
}

/** 範囲 r が [lo, hi] と重なる (端で接する場合を含む) か */
function touches(r: RevisionRange, lo: number, hi: number): boolean {
  return r.end.docOffset >= lo && r.start.docOffset <= hi;
}

/**
 * 位置付きの編集イベント列から、挿入の窓と特徴量を求める。
 * 窓は開始時刻の順に並べる。挿入が1件だけの窓も含む。
 */
export function detectInsertionWindows(
  events: PositionedRevisionEvent[],
  options: WindowOptions = DEFAULT_WINDOW_OPTIONS
): InsertionWindowResult {
  const dated = events.filter((e) => e.date);
  const timeResolutionSec = estimateTimeResolutionSec(dated.map((e) => e.date!));
  const cands = windowCandidates(events).sort((a, b) => a.date!.getTime() - b.date!.getTime());
  const maxDtMs = options.seconds * 1000;

  // 時刻順に見ていき、開始から seconds 以内で、文書上の距離が近い窓に加える。無ければ新しい窓を開く
  interface Open {
    start: number;
    range: RevisionRange;
    members: PositionedRevisionEvent[];
  }
  const groups: Open[] = [];
  let firstOpen = 0;
  for (const e of cands) {
    const t = e.date!.getTime();
    while (firstOpen < groups.length && t - groups[firstOpen].start > maxDtMs) firstOpen++;
    const g = groups
      .slice(firstOpen)
      .find(
        (g) =>
          t - g.start <= maxDtMs &&
          (!options.byAuthor || g.members[0].author === e.author) &&
          (options.chars === undefined || rangeGap(g.range, e.final) <= options.chars) &&
          (options.paras === undefined || paraGap(g.range, e.final) <= options.paras)
      );
    if (g) {
      g.members.push(e);
      g.range = {
        start: e.final.start.docOffset < g.range.start.docOffset ? e.final.start : g.range.start,
        end: e.final.end.docOffset > g.range.end.docOffset ? e.final.end : g.range.end,
      };
    } else {
      groups.push({ start: t, range: e.final, members: [e] });
    }
  }

  const others = events.filter((e) => !e.move && e.date);
  const windows: InsertionWindow[] = [];
  for (const { members, range } of groups) {
    const memberIds = new Set(members.map((e) => e.id));
    const times = members.map((e) => e.date!.getTime());
    const start = Math.min(...times);
    const end = Math.max(...times);
    const { start: rangeStart, end: rangeEnd } = range;
    const lo = rangeStart.docOffset;
    const hi = rangeEnd.docOffset;

    const insertedChars = members.reduce((n, e) => n + e.chars, 0);
    const durationSec = (end - start) / 1000 + Math.max(MIN_DURATION_SEC, timeResolutionSec);

    let postEditChars = 0;
    let precedingDeletedChars = 0;
    for (const e of others) {
      if (e.type === "ins" && memberIds.has(e.id)) continue;
      if (!touches(e.final, lo, hi)) continue;
      const t = e.date!.getTime();
      if (t > end) postEditChars += e.chars;
      else if (e.type === "del" && t >= start - maxDtMs) precedingDeletedChars += e.chars;
    }

    windows.push({
      index: 0,
      start: new Date(start),
      end: new Date(end),
      authors: [...new Set(members.map((e) => e.author).filter((a): a is string => !!a))],
      eventIds: members.map((e) => e.id),
      range: { start: rangeStart, end: rangeEnd },
      features: {
        insertedChars,
        durationSec,
        cps: insertedChars / durationSec,
        spanChars: hi - lo,
        spanParas: rangeEnd.paraIndex - rangeStart.paraIndex + 1,
        maxSingleInsert: Math.max(...members.map((e) => e.chars)),
        paraCount: new Set(members.flatMap((e) => e.paraModelIndices)).size,
        postEditRatio: postEditChars / insertedChars,
        precededByDeletion: precedingDeletedChars > 0,
        precedingDeletedChars,
        insertCount: members.length,
      },
    });
  }
  windows.sort((a, b) => a.start.getTime() - b.start.getTime() || a.range.start.docOffset - b.range.start.docOffset);
  windows.forEach((w, i) => (w.index = i));
  return { options, timeResolutionSec, windows };
}

/** JSON 出力用の窓 (日時は ISO 8601 文字列) */
export interface InsertionWindowJson {
  index: number;
  start: string;
  end: string;
  authors: string[];
  eventIds: string[];
  range: {
    startParaIndex: number;
    startOffsetInPara: number;
    startDocOffset: number;
    endParaIndex: number;
    endOffsetInPara: number;
    endDocOffset: number;
  };
  features: WindowFeatures;
}

/** 窓の検出結果を JSON にできる形に変換する */
export function insertionWindowsToJson(result: InsertionWindowResult): {
  window: WindowOptions;
  timeResolutionSec: number;
  windows: InsertionWindowJson[];
} {
  return {
    window: result.options,
    timeResolutionSec: result.timeResolutionSec,
    windows: result.windows.map((w) => ({
      index: w.index,
      start: w.start.toISOString(),
      end: w.end.toISOString(),
      authors: w.authors,
      eventIds: w.eventIds,
      range: {
        startParaIndex: w.range.start.paraIndex,
        startOffsetInPara: w.range.start.offsetInPara,
        startDocOffset: w.range.start.docOffset,
        endParaIndex: w.range.end.paraIndex,
        endOffsetInPara: w.range.end.offsetInPara,
        endDocOffset: w.range.end.docOffset,
      },
      features: w.features,
    })),
  };
}
