/**
 * insertionRules.ts
 *
 * 挿入の窓 (insertionWindows.ts) の特徴量に対する条件で、窓に段階 (レベル) を付ける判定ルール。
 * ルールはルールファイル (YAML / JSON) に書く。
 *
 *   ruleSet: example-v1        # 出力に記録する識別子
 *   window: { seconds: 60, chars: 2000 }
 *   levels:                    # 上から評価し、最初に当てはまったレベルを採用
 *     - id: level-2
 *       label: { ja: 貼り付けの疑い (大), en: Likely paste (large) }
 *       color: "#C2410C"
 *       when: { all: [ { insertedChars: { gte: 800 } }, { postEditRatio: { lt: 0.05 } } ] }
 *     - id: level-1
 *       label: { ja: 一括挿入文字数過多, en: Large bulk insertion }
 *       color: "#F28C28"
 *       when: { insertedChars: { gte: 300 } }
 *
 * 条件 (when):
 *  - { <特徴量>: { gte|gt|lte|lt|eq: 値, ... } }  1つのオブジェクトに複数書いた比較はすべて満たす必要がある
 *  - { <特徴量>: 値 }                              eq の省略形 (precededByDeletion: true など)
 *  - { all: [条件, ...] } / { any: [条件, ...] } / { not: 条件 }
 *
 * ルールファイルが無いときは、--bulk-chars から作る既定ルール (defaultRuleSet) を使う。
 * これは「同じ作成者・同じ日時の挿入の合計が bulkChars 文字以上なら一括挿入」という従来の規則と同じ。
 */

import type { FeatureName, InsertionWindowResult, WindowFeatures, WindowOptions } from "./insertionWindows";
import { DEFAULT_WINDOW_OPTIONS } from "./insertionWindows";
import { getLang, Lang, t } from "./i18n";

export type LocalizedText = string | Partial<Record<Lang, string>>;

export type CompareOp = "gte" | "gt" | "lte" | "lt" | "eq";

export type Condition =
  | { all: Condition[] }
  | { any: Condition[] }
  | { not: Condition }
  | { features: Partial<Record<FeatureName, Partial<Record<CompareOp, number | boolean>>>> };

export interface RuleLevel {
  id: string;
  label: LocalizedText;
  color: string;
  when: Condition;
}

export interface RuleSet {
  /** 出力に記録する識別子 */
  ruleSet: string;
  window: WindowOptions;
  /** 上から評価し、最初に当てはまったレベルを採用する */
  levels: RuleLevel[];
}

/** 一括挿入 (オレンジ) の色。図の既定の色と同じ */
export const DEFAULT_BULK_COLOR = "#EA6C00";

/** 既定ルールのレベル id。図ではこのレベルを従来の「一括挿入」として描く */
export const DEFAULT_BULK_LEVEL_ID = "bulk";

/** --bulk-chars に相当する既定ルール (同じ作成者・同じ日時の挿入の合計が bulkChars 文字以上) */
export function defaultRuleSet(bulkChars: number): RuleSet {
  return {
    ruleSet: `default(bulk-chars=${bulkChars})`,
    window: { seconds: 0, byAuthor: true },
    levels: [
      {
        id: DEFAULT_BULK_LEVEL_ID,
        label: { en: "Bulk insertion", ja: "一括挿入" },
        color: DEFAULT_BULK_COLOR,
        when: { features: { insertedChars: { gte: bulkChars } } },
      },
    ],
  };
}

const NUMERIC_FEATURES: FeatureName[] = [
  "insertedChars",
  "durationSec",
  "cps",
  "spanChars",
  "spanParas",
  "maxSingleInsert",
  "paraCount",
  "postEditRatio",
  "precedingDeletedChars",
  "insertCount",
];
const BOOLEAN_FEATURES: FeatureName[] = ["precededByDeletion"];
export const FEATURE_NAMES: readonly FeatureName[] = [...NUMERIC_FEATURES, ...BOOLEAN_FEATURES];
const OPS: CompareOp[] = ["gte", "gt", "lte", "lt", "eq"];

// ---------------------------------------------------------------------------
// 検証
// ---------------------------------------------------------------------------

/** ルールファイルの内容が不正なときのエラー (message は表示言語の文) */
export class RuleSetError extends Error {}

function fail(where: string, message: string): never {
  throw new RuleSetError(`${where}: ${message}`);
}

function isMapping(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function checkKeys(v: Record<string, unknown>, allowed: string[], where: string): void {
  for (const k of Object.keys(v)) {
    if (!allowed.includes(k)) fail(where, t("rulesUnknownKey", k, allowed.join(", ")));
  }
}

function num(v: unknown, where: string, min: number, allowZero: boolean): number {
  if (typeof v !== "number" || !Number.isFinite(v) || v < min || (!allowZero && v === 0)) {
    fail(where, t(allowZero ? "rulesNotNonNegative" : "rulesNotPositive"));
  }
  return v;
}

function parseCondition(v: unknown, where: string): Condition {
  if (!isMapping(v)) fail(where, t("rulesNotMapping"));
  const keys = Object.keys(v);
  if (keys.length === 0) fail(where, t("rulesEmptyCondition"));
  for (const k of ["all", "any"] as const) {
    if (k in v) {
      if (keys.length > 1) fail(where, t("rulesMixedCondition", k));
      const list = v[k];
      if (!Array.isArray(list) || list.length === 0) fail(`${where}.${k}`, t("rulesNotList"));
      const conds = list.map((c, i) => parseCondition(c, `${where}.${k}[${i}]`));
      return k === "all" ? { all: conds } : { any: conds };
    }
  }
  if ("not" in v) {
    if (keys.length > 1) fail(where, t("rulesMixedCondition", "not"));
    return { not: parseCondition(v.not, `${where}.not`) };
  }
  const features: Partial<Record<FeatureName, Partial<Record<CompareOp, number | boolean>>>> = {};
  for (const [name, spec] of Object.entries(v)) {
    const w = `${where}.${name}`;
    if (!FEATURE_NAMES.includes(name as FeatureName)) {
      fail(w, t("rulesUnknownFeature", name, FEATURE_NAMES.join(", ")));
    }
    const feature = name as FeatureName;
    const isBool = BOOLEAN_FEATURES.includes(feature);
    // { feature: 値 } は eq の省略形
    const ops = isMapping(spec) ? spec : { eq: spec };
    if (Object.keys(ops).length === 0) fail(w, t("rulesEmptyCondition"));
    const out: Partial<Record<CompareOp, number | boolean>> = {};
    for (const [op, value] of Object.entries(ops)) {
      if (!OPS.includes(op as CompareOp)) fail(w, t("rulesUnknownOp", op, OPS.join(", ")));
      if (isBool) {
        if (op !== "eq") fail(`${w}.${op}`, t("rulesBooleanOnlyEq"));
        if (typeof value !== "boolean") fail(`${w}.${op}`, t("rulesNotBoolean"));
      } else if (typeof value !== "number" || !Number.isFinite(value)) {
        fail(`${w}.${op}`, t("rulesNotNumber"));
      }
      out[op as CompareOp] = value as number | boolean;
    }
    features[feature] = out;
  }
  return { features };
}

function parseLabel(v: unknown, where: string): LocalizedText {
  if (typeof v === "string" && v.trim() !== "") return v;
  if (isMapping(v)) {
    checkKeys(v, ["ja", "en"], where);
    const out: Partial<Record<Lang, string>> = {};
    for (const [k, s] of Object.entries(v)) {
      if (typeof s !== "string" || s.trim() === "") fail(`${where}.${k}`, t("rulesNotString"));
      out[k as Lang] = s;
    }
    if (Object.keys(out).length > 0) return out;
  }
  return fail(where, t("rulesBadLabel"));
}

function parseWindow(v: unknown, where: string): WindowOptions {
  if (v === undefined) return { ...DEFAULT_WINDOW_OPTIONS };
  if (!isMapping(v)) fail(where, t("rulesNotMapping"));
  checkKeys(v, ["seconds", "chars", "paras", "byAuthor"], where);
  if (v.seconds === undefined) fail(`${where}.seconds`, t("rulesMissing"));
  const w: WindowOptions = { seconds: num(v.seconds, `${where}.seconds`, 0, true) };
  if (v.chars !== undefined) w.chars = num(v.chars, `${where}.chars`, 0, true);
  if (v.paras !== undefined) w.paras = num(v.paras, `${where}.paras`, 0, true);
  if (v.byAuthor !== undefined) {
    if (typeof v.byAuthor !== "boolean") fail(`${where}.byAuthor`, t("rulesNotBoolean"));
    w.byAuthor = v.byAuthor;
  }
  return w;
}

/** ルールファイルを読み込んだ値 (YAML / JSON を解析したもの) を検証してルールにする */
export function parseRuleSet(value: unknown): RuleSet {
  if (!isMapping(value)) fail("(root)", t("rulesNotMapping"));
  checkKeys(value, ["ruleSet", "window", "levels"], "(root)");
  if (typeof value.ruleSet !== "string" || value.ruleSet.trim() === "") fail("ruleSet", t("rulesNotString"));
  const levelsRaw = value.levels;
  if (!Array.isArray(levelsRaw) || levelsRaw.length === 0) fail("levels", t("rulesNotList"));
  const seen = new Set<string>();
  const levels = levelsRaw.map((lv, i): RuleLevel => {
    const where = `levels[${i}]`;
    if (!isMapping(lv)) fail(where, t("rulesNotMapping"));
    checkKeys(lv, ["id", "label", "color", "when"], where);
    if (typeof lv.id !== "string" || lv.id.trim() === "") fail(`${where}.id`, t("rulesNotString"));
    if (seen.has(lv.id)) fail(`${where}.id`, t("rulesDuplicateLevel", lv.id));
    seen.add(lv.id);
    if (typeof lv.color !== "string" || !/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(lv.color)) {
      fail(`${where}.color`, t("rulesBadColor"));
    }
    if (lv.when === undefined) fail(`${where}.when`, t("rulesMissing"));
    return {
      id: lv.id,
      label: parseLabel(lv.label ?? lv.id, `${where}.label`),
      color: lv.color,
      when: parseCondition(lv.when, `${where}.when`),
    };
  });
  return { ruleSet: value.ruleSet, window: parseWindow(value.window, "window"), levels };
}

// ---------------------------------------------------------------------------
// 評価
// ---------------------------------------------------------------------------

function compare(actual: number | boolean, op: CompareOp, expected: number | boolean): boolean {
  if (op === "eq") return actual === expected;
  const a = Number(actual);
  const e = Number(expected);
  return op === "gte" ? a >= e : op === "gt" ? a > e : op === "lte" ? a <= e : a < e;
}

/** 特徴量が条件を満たすか */
export function evaluateCondition(cond: Condition, f: WindowFeatures): boolean {
  if ("all" in cond) return cond.all.every((c) => evaluateCondition(c, f));
  if ("any" in cond) return cond.any.some((c) => evaluateCondition(c, f));
  if ("not" in cond) return !evaluateCondition(cond.not, f);
  return Object.entries(cond.features).every(([name, ops]) =>
    Object.entries(ops ?? {}).every(([op, v]) => compare(f[name as FeatureName], op as CompareOp, v!))
  );
}

/** 窓ごとに、最初に当てはまったレベル (無ければ undefined) を返す (result.windows と同じ順) */
export function assignLevels(rules: RuleSet, result: InsertionWindowResult): (RuleLevel | undefined)[] {
  return result.windows.map((w) => rules.levels.find((lv) => evaluateCondition(lv.when, w.features)));
}

/** 挿入の w:id → その挿入を含む窓のレベル (レベルの付いた窓に含まれる挿入だけ) */
export function levelsByInsertion(
  result: InsertionWindowResult,
  levels: (RuleLevel | undefined)[]
): Map<string, RuleLevel> {
  const map = new Map<string, RuleLevel>();
  result.windows.forEach((w, i) => {
    const lv = levels[i];
    if (lv) for (const id of w.eventIds) map.set(id, lv);
  });
  return map;
}

/** レベルの表示名 (表示言語の名前が無ければ、もう一方の言語の名前) */
export function levelLabel(level: RuleLevel, lang: Lang = getLang()): string {
  const l = level.label;
  if (typeof l === "string") return l;
  return l[lang] ?? l.en ?? l.ja ?? level.id;
}
