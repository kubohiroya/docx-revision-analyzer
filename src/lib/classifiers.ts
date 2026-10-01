/**
 * classifiers.ts
 *
 * 挿入の分類を「分類器のパイプライン」として行う。
 * 分類器は解析結果 (位置付きのイベント列・挿入の窓と特徴量・時間区間) を受け取り、
 * ハイライト (カテゴリ id・対象の挿入・時間範囲・文書範囲・特徴量・理由の文章) を返す。
 * 複数の分類器の結果は、挿入ごとに最も priority の高いカテゴリ (同じなら先に並んだ分類器) にまとめ、
 * categories.ts のカテゴリで描画する。
 *
 * 既定の分類器:
 *  - builtin.relocation: 並べ替え・複製 (挿入された文章と同じ内容が文書のどこかで削除されている)。
 *    カテゴリは移動 (moved)。insertionKinds.ts の規則と同じ
 *  - builtin.rules: 判定ルール (insertionRules.ts) のレベルの付いた窓の挿入。カテゴリはレベルごと
 *    (既定ルールなら既定の一括挿入 bulk)
 */

import type { DocxLayoutModel } from "./docxLayout";
import type { RevisionEvent } from "./docxRevisions";
import type { PositionedRevisionEvent, RevisionRange } from "./revisionPositions";
import { detectInsertionWindows, InsertionWindowResult, WindowOptions } from "./insertionWindows";
import { matchesCorpus, normalizeForMatch } from "./insertionKinds";
import {
  assignLevels,
  categoriesFromRules,
  categoryIdOfLevel,
  levelLabel,
  RuleLevel,
  RuleSet,
} from "./insertionRules";
import { builtinCategories, Category, CATEGORY_IDS, CategoryRegistry, LocalizedText } from "./categories";

/** 分類器に渡す解析結果 */
export interface AnalysisContext {
  model: DocxLayoutModel;
  /** 位置付きの編集イベント列 (revisionPositions.ts) */
  positioned: PositionedRevisionEvent[];
  /** 窓の設定ごとに挿入の窓と特徴量を求める (同じ設定の結果は使い回す) */
  windowsFor(options: WindowOptions): InsertionWindowResult;
  /** 無編集期間で分けた時間区間 (docx-revision-flow の区間と同じ分け方) */
  sessions: { start: Date; end: Date }[];
}

export interface Highlight {
  /** このハイライトを返した分類器の id */
  classifierId: string;
  /** 描画に使うカテゴリの id (categories.ts) */
  categoryId: string;
  /** 対象の挿入の w:id */
  eventIds: string[];
  timeRange: { start: Date; end: Date };
  /** 最終文書での範囲 */
  docRange?: RevisionRange;
  /** 判定に使った特徴量 */
  features?: Record<string, number | boolean>;
  /** 判定の理由 (利用者に見せる文章) */
  reason: LocalizedText;
  /** 判定ルールのレベル id (builtin.rules のとき) */
  level?: string;
}

export interface Classifier {
  id: string;
  version: string;
  /** この分類器が使うカテゴリ (登録してから描画する) */
  categories?: Category[];
  classify(ctx: AnalysisContext): Highlight[] | Promise<Highlight[]>;
}

/** 位置付きイベント列から解析結果を作る */
export function createAnalysisContext(
  model: DocxLayoutModel,
  positioned: PositionedRevisionEvent[],
  gapThresholdHours: number
): AnalysisContext {
  const cache = new Map<string, InsertionWindowResult>();
  const dated = positioned
    .filter((e) => e.date)
    .map((e) => e.date!.getTime())
    .sort((a, b) => a - b);
  const sessions: { start: Date; end: Date }[] = [];
  const gapMs = Math.max(0, gapThresholdHours) * 3600 * 1000;
  for (const t of dated) {
    const last = sessions[sessions.length - 1];
    if (last && t - last.end.getTime() <= gapMs) last.end = new Date(t);
    else sessions.push({ start: new Date(t), end: new Date(t) });
  }
  return {
    model,
    positioned,
    sessions,
    windowsFor(options) {
      const key = JSON.stringify(options);
      let r = cache.get(key);
      if (!r) {
        r = detectInsertionWindows(positioned, options);
        cache.set(key, r);
      }
      return r;
    },
  };
}

// ---------------------------------------------------------------------------
// 既定の分類器
// ---------------------------------------------------------------------------

/** 並べ替え・複製: 挿入された文章と同じ内容が文書のどこかで削除されている */
export function relocationClassifier(): Classifier {
  return {
    id: "builtin.relocation",
    version: "1.0.0",
    classify(ctx) {
      const delCorpus = ctx.positioned
        .filter((e) => e.type === "del" && !e.move && e.date)
        .map((e) => normalizeForMatch(e.text))
        .join("");
      return ctx.positioned
        .filter((e) => e.type === "ins" && !e.move && e.date && matchesCorpus(e.text, delCorpus))
        .map((e) => ({
          classifierId: "builtin.relocation",
          categoryId: CATEGORY_IDS.moved,
          eventIds: [e.id],
          timeRange: { start: e.date!, end: e.date! },
          docRange: e.final,
          features: { chars: e.chars },
          reason: {
            en: "The same text was deleted elsewhere in the document (reordering or duplication)",
            ja: "同じ内容の文章が文書のどこかで削除されている (並べ替え・複製)",
          },
        }));
    },
  };
}

function fmtNum(n: number): string {
  return Number.isInteger(n) ? String(n) : n.toFixed(2);
}

function ruleReason(level: RuleLevel, f: { insertedChars: number; durationSec: number; paraCount: number }): LocalizedText {
  return {
    en: `${levelLabel(level, "en")}: ${f.insertedChars} chars inserted within ${fmtNum(f.durationSec)} s across ${f.paraCount} paragraph(s)`,
    ja: `${levelLabel(level, "ja")}: ${fmtNum(f.durationSec)} 秒のうちに ${f.paraCount} 段落へ ${f.insertedChars} 文字を挿入`,
  };
}

/** 判定ルールのレベルの付いた窓の挿入 */
export function ruleClassifier(rules: RuleSet): Classifier {
  return {
    id: "builtin.rules",
    version: "1.0.0",
    categories: categoriesFromRules(rules),
    classify(ctx) {
      const windows = ctx.windowsFor(rules.window);
      const levels = assignLevels(rules, windows);
      const out: Highlight[] = [];
      windows.windows.forEach((w, i) => {
        const lv = levels[i];
        if (!lv) return;
        out.push({
          classifierId: "builtin.rules",
          categoryId: categoryIdOfLevel(lv),
          eventIds: w.eventIds,
          timeRange: { start: w.start, end: w.end },
          docRange: w.range,
          features: { ...w.features },
          reason: ruleReason(lv, w.features),
          level: lv.id,
        });
      });
      return out;
    },
  };
}

/** 既定の分類器のパイプライン (並べ替え → 判定ルール) */
export function defaultClassifiers(rules: RuleSet): Classifier[] {
  return [relocationClassifier(), ruleClassifier(rules)];
}

// ---------------------------------------------------------------------------
// パイプラインの実行とまとめ
// ---------------------------------------------------------------------------

export interface ClassificationResult {
  /** 実行した分類器 (出力に記録する) */
  classifiers: { id: string; version: string }[];
  /** すべての分類器のハイライト (分類器の順) */
  highlights: Highlight[];
  /** 描画に使うカテゴリ (既定のカテゴリ + 分類器のカテゴリ) */
  categories: CategoryRegistry;
  /** 挿入の w:id → その挿入に採用したハイライト */
  byInsertion: Map<string, Highlight>;
}

/**
 * 分類器を順に実行し、結果をまとめる。
 * カテゴリは既定のカテゴリから一括挿入 (bulk) を除いたものに、各分類器のカテゴリを登録して作る
 * (判定ルールのレベルが一括挿入を置き換える。既定ルールは一括挿入を登録し直す)。
 * 挿入ごとに、カテゴリの priority が最も高いハイライト (同じなら先の分類器のもの) を採用する。
 */
export async function runClassifiers(classifiers: Classifier[], ctx: AnalysisContext): Promise<ClassificationResult> {
  const categories = new CategoryRegistry(builtinCategories().filter((c) => c.id !== CATEGORY_IDS.bulk));
  for (const c of classifiers) for (const cat of c.categories ?? []) categories.register(cat);

  const highlights: Highlight[] = [];
  for (const c of classifiers) highlights.push(...(await c.classify(ctx)));

  const priorityOf = (h: Highlight) => categories.get(h.categoryId)?.priority ?? 0;
  const byInsertion = new Map<string, Highlight>();
  for (const h of highlights) {
    for (const id of h.eventIds) {
      const cur = byInsertion.get(id);
      if (!cur || priorityOf(h) > priorityOf(cur)) byInsertion.set(id, h);
    }
  }
  return {
    classifiers: classifiers.map((c) => ({ id: c.id, version: c.version })),
    highlights,
    categories,
    byInsertion,
  };
}

/** JSON 出力用のハイライト (日時は ISO 8601 文字列) */
export function highlightToJson(h: Highlight): Record<string, unknown> {
  return {
    classifier: h.classifierId,
    category: h.categoryId,
    level: h.level ?? null,
    eventIds: h.eventIds,
    start: h.timeRange.start.toISOString(),
    end: h.timeRange.end.toISOString(),
    docRange: h.docRange
      ? { startDocOffset: h.docRange.start.docOffset, endDocOffset: h.docRange.end.docOffset }
      : null,
    features: h.features ?? null,
    reason: h.reason,
  };
}

/**
 * まとめた結果で RevisionEvent の insKind / category / level を設定する (docx-revision-chart 用)。
 * 移動のカテゴリなら moved、ハイライトのカテゴリなら bulk、どれにも当たらなければ fine。
 */
export function applyClassification(events: RevisionEvent[], result: ClassificationResult): void {
  for (const e of events) {
    if (e.type !== "ins") continue;
    const h = result.byInsertion.get(String(e.id));
    const cat = h ? result.categories.get(h.categoryId) : undefined;
    e.insKind = cat?.role === "moved" ? "moved" : cat?.role === "highlight" ? "bulk" : "fine";
    e.category = cat?.role === "highlight" ? cat.id : undefined;
    e.level = cat?.role === "highlight" ? h!.level : undefined;
  }
}

/** 挿入の w:id → ハイライトのカテゴリの id (FlowOptions.highlightOf 用。移動などは含めない) */
export function highlightCategoryMap(result: ClassificationResult): Map<string, string> {
  const map = new Map<string, string>();
  for (const [id, h] of result.byInsertion) {
    if (result.categories.get(h.categoryId)?.role === "highlight") map.set(id, h.categoryId);
  }
  return map;
}
