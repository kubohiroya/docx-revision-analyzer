import { RevisionEvent } from "./docxRevisions";

/**
 * 「AI不正利用疑いスコア」算出ロジック
 * ------------------------------------------------------------
 * 目的:
 *   WordのGUIで自分の手でタイプ・校閲したのではなく、外部アプリ(ChatGPT等)で
 *   作文した完成文をコピー&ペーストして仕上げたと推測できるような、
 *   「不自然に極端な文字数の増加」を検出し、0(疑いなし)〜100(疑い濃厚)の
 *   スコアとして数値化する。
 *
 * 前提と限界 (README.md にも明記):
 *   - これは統計的なヒューリスティックであり、不正の「証拠」にはならない。
 *     早いタイピスト、音声入力、下書きを別ファイルで書いてから貼り付けた
 *     正当なケースなどでもスコアが高くなり得る (誤検知の可能性)。
 *   - w:date のタイムスタンプは秒単位精度のため、1秒未満の間隔は区別できない。
 *   - Word は連続した通常のタイピングであっても、保存タイミングや編集の区切りに
 *     よって複数の挿入イベントにまとまることがあり、逆に短時間の入力が1つの
 *     w:ins にまとまることもある。本ロジックはその挙動のばらつきを完全には
 *     補正できない。
 *
 * 考え方:
 *   各「挿入イベント」(w:ins 1個) について、
 *     (a) バーストスコア: そのイベント単体の文字数の大きさ
 *         (burstLowChars 文字以下なら0、burstHighChars 文字以上なら100)
 *     (b) 速度スコア: 直前の挿入イベントからの経過時間に対する挿入速度
 *         (rateLowCps 文字/秒以下なら0、rateHighCps 文字/秒以上なら100)
 *   の大きい方を採用し、そのイベントの疑いスコアとする
 *   (どちらか一方が極端でも「怪しい」と判断するため max を取る)。
 *   ただし、あまりに小さい挿入 (minCharsToFlag 文字未満) は
 *   誤検知防止のため常にスコア0とする。
 *
 *   文書全体のスコアは、
 *     0.6 * (最も疑わしいイベントのスコア)
 *   + 0.4 * (文字数で重み付けした平均スコア)
 *   の加重平均とする。最大値だけで決めると1回の誤検知に引っ張られやすく、
 *   平均だけで決めると大きな1回のペーストが他の多数の小さな正常な編集で
 *   薄まってしまうため、両者をブレンドしている。
 */

export interface SuspicionOptions {
  /** これ未満の文字数の挿入イベントは常にスコア0とする (誤検知防止) */
  minCharsToFlag: number;
  /** バーストスコアが0になる文字数境界 */
  burstLowChars: number;
  /** バーストスコアが100になる文字数境界 */
  burstHighChars: number;
  /** 速度スコアが0になる境界 (文字/秒) */
  rateLowCps: number;
  /** 速度スコアが100になる境界 (文字/秒) */
  rateHighCps: number;
  /** 文書全体スコアにおける「最大イベントスコア」の重み (残りは文字数加重平均) */
  maxWeight: number;
}

export const DEFAULT_SUSPICION_OPTIONS: SuspicionOptions = {
  minCharsToFlag: 20,
  burstLowChars: 20,
  burstHighChars: 300,
  rateLowCps: 8,
  rateHighCps: 40,
  maxWeight: 0.6,
};

export interface EventScore {
  index: number;
  date: string;
  author: string | undefined;
  chars: number;
  dtSeconds: number | null;
  impliedCps: number | null;
  burstScore: number;
  rateScore: number;
  eventScore: number;
}

export interface SuspicionReport {
  score: number;
  riskLevel: "low" | "medium" | "high" | "very_high";
  options: SuspicionOptions;
  totalInsertedChars: number;
  totalDeletedChars: number;
  insertionEventCount: number;
  topSuspiciousEvents: EventScore[];
  allEventScores: EventScore[];
  notes: string[];
}

function clamp(x: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, x));
}

function scale(value: number, lo: number, hi: number): number {
  if (hi <= lo) return value >= hi ? 100 : 0;
  return clamp(((value - lo) / (hi - lo)) * 100, 0, 100);
}

function classify(score: number): SuspicionReport["riskLevel"] {
  if (score >= 75) return "very_high";
  if (score >= 45) return "high";
  if (score >= 20) return "medium";
  return "low";
}

export function computeSuspicionScore(
  events: RevisionEvent[],
  optionsIn: Partial<SuspicionOptions> = {}
): SuspicionReport {
  const options: SuspicionOptions = { ...DEFAULT_SUSPICION_OPTIONS, ...optionsIn };
  const notes: string[] = [];

  const insertions = events
    .filter((e) => e.type === "ins")
    .sort((a, b) => a.date.getTime() - b.date.getTime());

  const totalInsertedChars = insertions.reduce((s, e) => s + e.chars, 0);
  const totalDeletedChars = events
    .filter((e) => e.type === "del")
    .reduce((s, e) => s + e.chars, 0);

  if (insertions.length === 0) {
    notes.push("挿入イベントが見つからなかったため、スコアは算出できません (0を返します)。");
    return {
      score: 0,
      riskLevel: "low",
      options,
      totalInsertedChars,
      totalDeletedChars,
      insertionEventCount: 0,
      topSuspiciousEvents: [],
      allEventScores: [],
      notes,
    };
  }

  const allEventScores: EventScore[] = [];
  let prevDate: Date | null = null;

  insertions.forEach((ev, idx) => {
    let dtSeconds: number | null = null;
    let impliedCps: number | null = null;
    if (prevDate) {
      dtSeconds = Math.max(1, (ev.date.getTime() - prevDate.getTime()) / 1000);
      impliedCps = ev.chars / dtSeconds;
    }
    prevDate = ev.date;

    let burstScore = scale(ev.chars, options.burstLowChars, options.burstHighChars);
    let rateScore = impliedCps === null ? 0 : scale(impliedCps, options.rateLowCps, options.rateHighCps);

    let eventScore = Math.max(burstScore, rateScore);
    if (ev.chars < options.minCharsToFlag) {
      eventScore = 0;
      burstScore = 0;
      rateScore = 0;
    }

    allEventScores.push({
      index: idx,
      date: ev.date.toISOString(),
      author: ev.author,
      chars: ev.chars,
      dtSeconds,
      impliedCps,
      burstScore: Math.round(burstScore * 10) / 10,
      rateScore: Math.round(rateScore * 10) / 10,
      eventScore: Math.round(eventScore * 10) / 10,
    });
  });

  const maxEventScore = Math.max(...allEventScores.map((e) => e.eventScore));
  const weightedAvg =
    allEventScores.reduce((s, e) => s + e.eventScore * e.chars, 0) /
    Math.max(1, allEventScores.reduce((s, e) => s + e.chars, 0));

  const rawScore =
    options.maxWeight * maxEventScore + (1 - options.maxWeight) * weightedAvg;
  const score = Math.round(clamp(rawScore, 0, 100));

  const topSuspiciousEvents = [...allEventScores]
    .sort((a, b) => b.eventScore - a.eventScore)
    .slice(0, 5)
    .filter((e) => e.eventScore > 0);

  if (insertions.length < 5) {
    notes.push(
      "挿入イベント数が少ないため (5件未満)、統計的な信頼性は低くなります。参考値として扱ってください。"
    );
  }

  return {
    score,
    riskLevel: classify(score),
    options,
    totalInsertedChars,
    totalDeletedChars,
    insertionEventCount: insertions.length,
    topSuspiciousEvents,
    allEventScores,
    notes,
  };
}
