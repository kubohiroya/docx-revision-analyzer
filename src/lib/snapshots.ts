/**
 * snapshots.ts
 *
 * 同じ文書を何回かに分けて提出させたもの (スナップショット) を、順に並べて1本の編集履歴にする (通し解析)。
 *
 * 想定する使い方: 卒論のように長く書く文書では、OneDrive のバージョン履歴は古いものから消え、自分の未確定の
 * 挿入を書き直すと変更履歴にも残らない。そこで区切り (スプリント) ごとに提出させて教員が保管し、変更を
 * すべて承諾してから返す。承諾した後の文章を書き直すと削除として記録されるため、区切りをまたぐ推敲は残る。
 *
 * 各回の提出物には、前回までの変更が残っていることも (承諾しなかった場合)、残っていないこともある (承諾した場合)。
 * どちらでも同じ変更を二重に数えないよう、種類・作成者・日時・文章が同じ変更は1つとみなす。
 *
 * ファイルシステムは使わない (保管は node/snapshotArchive.ts)。
 */

import type { DocxRevisionData, RevisionEvent } from "./docxRevisions";
import type { Session } from "./sessions";
import type { TamperEvidence } from "./tamperEvidence";

/** 前回の提出より前の日時かを判定するときの余裕 (時計のずれ) */
const DATED_BEFORE_TOLERANCE_MS = 60 * 1000;

export interface SnapshotDelta {
  /** この回で初めて現れた変更 (時系列順) */
  newEvents: RevisionEvent[];
  /** 前回までにも現れていた変更の数 (承諾せずに返した場合に残っているもの) */
  carriedEvents: number;
  /** この回に挿入・削除した文字数 */
  inserted: number;
  deleted: number;
  /** この回で初めて現れたのに、前回までの最後の変更より前の日時の変更の数 */
  datedBeforePrevious: number;
  /** 前回までの最後の変更の日時 */
  previousLastDate?: Date;
}

export interface MergedSnapshots {
  deltas: SnapshotDelta[];
  /** 最初の回の、変更履歴の記録を始める前からあった文字数 */
  baselineCharCount: number;
}

/** 同じ変更とみなすためのキー */
export function revisionEventKey(e: RevisionEvent): string {
  return `${e.type}|${e.author ?? ""}|${e.date.toISOString()}|${e.text}`;
}

/** 提出順に並べた各回の変更履歴から、各回で新しく加わった変更を取り出す */
export function mergeSnapshots(snapshots: DocxRevisionData[]): MergedSnapshots {
  const seen = new Map<string, number>();
  let lastDate: Date | undefined;
  const deltas: SnapshotDelta[] = [];
  for (const data of snapshots) {
    // 同じキーの変更が同じ回に複数あることもあるため、回数で比べる
    const here = new Map<string, number>();
    const newEvents: RevisionEvent[] = [];
    for (const e of data.events) {
      const k = revisionEventKey(e);
      const n = (here.get(k) ?? 0) + 1;
      here.set(k, n);
      if (n > (seen.get(k) ?? 0)) newEvents.push(e);
    }
    for (const [k, n] of here) seen.set(k, Math.max(seen.get(k) ?? 0, n));
    newEvents.sort((a, b) => a.date.getTime() - b.date.getTime());
    const datedBeforePrevious = lastDate
      ? newEvents.filter((e) => e.date.getTime() < lastDate!.getTime() - DATED_BEFORE_TOLERANCE_MS).length
      : 0;
    deltas.push({
      newEvents,
      carriedEvents: data.events.length - newEvents.length,
      inserted: newEvents.filter((e) => e.type === "ins").reduce((s, e) => s + e.chars, 0),
      deleted: newEvents.filter((e) => e.type === "del").reduce((s, e) => s + e.chars, 0),
      datedBeforePrevious,
      previousLastDate: lastDate,
    });
    for (const e of newEvents) if (!lastDate || e.date > lastDate) lastDate = e.date;
  }
  return { deltas, baselineCharCount: snapshots[0]?.baselineCharCount ?? 0 };
}

/**
 * 通し解析のチャート (renderSessionedRevisionChart) に渡すセッション。1回の提出を1つのパネルにする。
 * 新しい変更の無い回は除き、indices に元の回の番号 (0 始まり) を返す
 */
export function snapshotSessions(merged: MergedSnapshots): { sessions: Session[]; indices: number[] } {
  const sessions: Session[] = [];
  const indices: number[] = [];
  let prevEnd: Date | undefined;
  merged.deltas.forEach((d, i) => {
    if (d.newEvents.length === 0) return;
    const first = d.newEvents[0].date;
    sessions.push({
      events: d.newEvents,
      gapBeforeHours: prevEnd ? Math.max(0, (first.getTime() - prevEnd.getTime()) / 3600000) : null,
    });
    indices.push(i);
    prevEnd = d.newEvents[d.newEvents.length - 1].date;
  });
  return { sessions, indices };
}

/** 前回の提出より前の日時の変更が今回初めて現れた、という痕跡 */
export function datedBeforePreviousEvidence(d: SnapshotDelta): TamperEvidence | undefined {
  if (d.datedBeforePrevious === 0 || !d.previousLastDate) return undefined;
  const when = d.previousLastDate.toISOString();
  return {
    id: "datedBeforePrevious",
    details: { changes: d.datedBeforePrevious, previousLastChange: when },
    message: {
      en:
        `${d.datedBeforePrevious} change(s) first seen in this submission are dated before the last change of the previous ` +
        `submission (${when}). The file may have been replaced with one edited elsewhere, or the dates rewritten.`,
      ja:
        `今回の提出で初めて現れた変更のうち ${d.datedBeforePrevious} 件が、前回の提出の最後の変更 (${when}) より前の日時です。` +
        "別の場所で編集したファイルに差し替えたか、日時が書き換えられた可能性があります。",
    },
  };
}
