import { RevisionEvent } from "./docxRevisions";

export interface Session {
  events: RevisionEvent[];
  /** このセッション開始前の無編集期間 (時間単位)。先頭セッションは null。 */
  gapBeforeHours: number | null;
}

/**
 * 変更履歴イベントを、「連続的に更新が行われた期間 (セッション)」ごとに分割する。
 *
 * イベントを時系列順に見ていき、直前のイベントからの経過時間が
 * thresholdHours (時間) を超えた場合、そこで新しいセッションに区切る。
 *
 * @param events 時系列順 (昇順) のイベント配列
 * @param thresholdHours 「更新がなかった期間」とみなす閾値 (時間)
 */
export function splitIntoSessions(
  events: RevisionEvent[],
  thresholdHours: number
): Session[] {
  if (events.length === 0) return [];

  const sorted = [...events].sort((a, b) => a.date.getTime() - b.date.getTime());
  const thresholdMs = Math.max(0, thresholdHours) * 3600 * 1000;

  const sessions: Session[] = [];
  let current: RevisionEvent[] = [sorted[0]];
  let currentGapBeforeHours: number | null = null;

  for (let i = 1; i < sorted.length; i++) {
    const prev = sorted[i - 1];
    const ev = sorted[i];
    const dtMs = ev.date.getTime() - prev.date.getTime();
    if (dtMs > thresholdMs) {
      sessions.push({ events: current, gapBeforeHours: currentGapBeforeHours });
      current = [ev];
      currentGapBeforeHours = dtMs / 3600000;
    } else {
      current.push(ev);
    }
  }
  sessions.push({ events: current, gapBeforeHours: currentGapBeforeHours });

  return sessions;
}
