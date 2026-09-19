import { RevisionEvent } from "./docxRevisions";

export interface Bucket {
  /** バケット開始時刻 */
  start: Date;
  /** バケット終了時刻 */
  end: Date;
  added: number;
  deleted: number;
  /** バケット終了時点での累計総文字数 */
  totalAtEnd: number;
}

export type BucketSpec = "auto" | "second" | "minute" | "hour" | "day" | number;

/** "auto" のとき、時間幅をおよそ target 個のバケットに分割する秒数を決める */
function resolveBucketSeconds(spanSeconds: number, spec: BucketSpec, target = 40): number {
  if (typeof spec === "number") return Math.max(1, spec);
  if (spec === "second") return 1;
  if (spec === "minute") return 60;
  if (spec === "hour") return 3600;
  if (spec === "day") return 86400;
  // auto: span を target 個程度に分割できる「キリのよい」秒数を選ぶ
  const candidates = [
    1, 5, 10, 30, 60, 300, 600, 900, 1800, 3600, 3 * 3600, 6 * 3600, 12 * 3600, 86400,
    7 * 86400, 30 * 86400,
  ];
  if (spanSeconds <= 0) return 60;
  const idealSeconds = spanSeconds / target;
  let chosen = candidates[candidates.length - 1];
  for (const c of candidates) {
    if (c >= idealSeconds) {
      chosen = c;
      break;
    }
  }
  return chosen;
}

/**
 * 挿入/削除イベントを時間バケットに集計し、各バケット終了時点での
 * 累計総文字数 (baseline + それまでの純増減) も併せて計算する。
 */
export function buildBuckets(
  events: RevisionEvent[],
  baselineCharCount: number,
  spec: BucketSpec = "auto"
): Bucket[] {
  if (events.length === 0) return [];

  const minTime = events[0].date.getTime();
  const maxTime = events[events.length - 1].date.getTime();
  const spanSeconds = Math.max(1, (maxTime - minTime) / 1000);
  const bucketSeconds = resolveBucketSeconds(spanSeconds, spec);
  const bucketMs = bucketSeconds * 1000;

  // バケット境界は minTime を起点に均等割り
  const bucketCount = Math.max(1, Math.ceil((maxTime - minTime + 1) / bucketMs));
  const buckets: Bucket[] = [];
  for (let i = 0; i < bucketCount; i++) {
    const start = new Date(minTime + i * bucketMs);
    const end = new Date(minTime + (i + 1) * bucketMs);
    buckets.push({ start, end, added: 0, deleted: 0, totalAtEnd: 0 });
  }

  let running = baselineCharCount;
  let bucketIdx = 0;
  for (const ev of events) {
    while (
      bucketIdx < buckets.length - 1 &&
      ev.date.getTime() >= buckets[bucketIdx].end.getTime()
    ) {
      // イベントが無いバケットも累計値を引き継ぐ
      buckets[bucketIdx].totalAtEnd = running;
      bucketIdx++;
    }
    if (ev.type === "ins") {
      buckets[bucketIdx].added += ev.chars;
      running += ev.chars;
    } else {
      buckets[bucketIdx].deleted += ev.chars;
      running -= ev.chars;
    }
  }
  // 残りのバケットに最終累計値を反映
  for (let i = bucketIdx; i < buckets.length; i++) {
    buckets[i].totalAtEnd = running;
  }

  return buckets;
}
