/**
 * categories.ts
 *
 * 図 (docx-revision-chart の積み上げ棒、docx-revision-flow の段落の塗り・帯) と凡例で使う
 * 色のカテゴリ。既定の5つ (細かい編集・一括挿入・移動・削除・変化なし) に加えて、
 * 判定ルールのレベル (insertionRules.ts) や拡張モジュールがカテゴリを登録できる。
 *
 * 役割 (role):
 *  - fine / moved / deleted / unchanged: 既定の分類。1つずつ
 *  - highlight: ハイライト (一括挿入や判定ルールのレベル)。いくつでも登録できる。
 *    既定の「一括挿入」(bulk) もこの役割で、判定ルールのレベルで置き換えると凡例から消える
 *
 * 1つの段落に複数のカテゴリが当てはまる場合は、priority の高い方で段落を塗り、2番目を左余白の
 * 細い帯で示す。priority が同じなら度合い (段落に占める割合など) の大きい方を優先する。
 * 既定のカテゴリはすべて priority 0 (従来どおり度合いだけで決まる)。
 *
 * 色だけに頼らないよう、カテゴリには模様 (pattern) を付けられる。模様は棒・段落の塗り・凡例に重ねて描く。
 */

import { getLang, Lang, t } from "./i18n";

export type LocalizedText = string | Partial<Record<Lang, string>>;

export type CategoryRole = "fine" | "highlight" | "moved" | "deleted" | "unchanged";

/** 色に重ねる模様 */
export type CategoryPattern = "hatch" | "cross" | "dots";

export const CATEGORY_PATTERNS: readonly CategoryPattern[] = ["hatch", "cross", "dots"];

export interface Category {
  id: string;
  role: CategoryRole;
  /** "#rrggbb" (または "#rgb") */
  color: string;
  /** 短い名前 (chart の凡例・ツールチップ)。関数なら表示のたびに呼ぶ (表示言語に合わせるため) */
  label: LocalizedText | (() => string);
  /** flow の凡例の説明。省略時は label */
  flowLegend?: LocalizedText | (() => string);
  priority: number;
  pattern?: CategoryPattern;
}

/** 既定カテゴリの id */
export const CATEGORY_IDS = {
  fine: "fine",
  bulk: "bulk",
  moved: "moved",
  deleted: "deleted",
  unchanged: "unchanged",
} as const;

export function builtinCategories(): Category[] {
  return [
    { id: CATEGORY_IDS.fine, role: "fine", color: "#16803D", priority: 0, label: () => t("kindFine"), flowLegend: () => t("flowLegendFine") },
    { id: CATEGORY_IDS.bulk, role: "highlight", color: "#EA6C00", priority: 0, label: () => t("kindBulk"), flowLegend: () => t("flowLegendBulk") },
    { id: CATEGORY_IDS.moved, role: "moved", color: "#2563EB", priority: 0, label: () => t("kindMoved"), flowLegend: () => t("flowLegendMoved") },
    { id: CATEGORY_IDS.deleted, role: "deleted", color: "#C62828", priority: 0, label: () => t("kindDeleted"), flowLegend: () => t("flowLegendDeleted") },
    { id: CATEGORY_IDS.unchanged, role: "unchanged", color: "#969BA5", priority: 0, label: () => t("flowLegendUnchanged") },
  ];
}

function resolveText(v: LocalizedText | (() => string), lang: Lang): string {
  if (typeof v === "function") return v();
  if (typeof v === "string") return v;
  return v[lang] ?? v.en ?? v.ja ?? "";
}

/** "#rrggbb" / "#rgb" → "r,g,b" */
export function hexToRgbTriple(hex: string): string {
  let h = hex.replace(/^#/, "");
  if (h.length === 3) h = [...h].map((c) => c + c).join("");
  const n = parseInt(h, 16);
  return `${(n >> 16) & 255},${(n >> 8) & 255},${n & 255}`;
}

/** 白背景とのコントラスト比 (WCAG 2.x の相対輝度による) */
export function contrastWithWhite(hex: string): number {
  const [r, g, b] = hexToRgbTriple(hex)
    .split(",")
    .map((v) => {
      const c = Number(v) / 255;
      return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
    });
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return 1.05 / (lum + 0.05);
}

/** 図の要素 (棒・凡例の色見本) が見分けられる最低限のコントラスト比 (WCAG 1.4.11) */
export const MIN_GRAPHIC_CONTRAST = 3;

export class CategoryRegistry {
  private readonly byId = new Map<string, Category>();

  constructor(categories: Category[] = builtinCategories()) {
    for (const c of categories) this.register(c);
  }

  /** カテゴリを登録する。同じ id があれば置き換える */
  register(c: Category): void {
    if (c.role !== "highlight") {
      // 既定の役割は1つずつ: 同じ役割の既存カテゴリを置き換える
      for (const [id, other] of this.byId) if (other.role === c.role) this.byId.delete(id);
    }
    this.byId.set(c.id, c);
  }

  unregister(id: string): void {
    this.byId.delete(id);
  }

  get(id: string): Category | undefined {
    return this.byId.get(id);
  }

  /** 役割のカテゴリ (fine / moved / deleted / unchanged) */
  byRole(role: Exclude<CategoryRole, "highlight">): Category {
    for (const c of this.byId.values()) if (c.role === role) return c;
    throw new Error(`no category for role ${role}`);
  }

  /** ハイライトのカテゴリ (priority の低い順。同じなら登録順) */
  highlights(): Category[] {
    return [...this.byId.values()].filter((c) => c.role === "highlight").sort((a, b) => a.priority - b.priority);
  }

  /** id のカテゴリ。無ければ既定の一括挿入、それも無ければ最初のハイライト */
  highlightOrDefault(id: string | undefined): Category {
    const c = (id !== undefined ? this.byId.get(id) : undefined) ?? this.byId.get(CATEGORY_IDS.bulk) ?? this.highlights()[0];
    if (!c) throw new Error("no highlight category");
    return c;
  }

  rgb(c: Category): string {
    return hexToRgbTriple(c.color);
  }

  label(c: Category, lang: Lang = getLang()): string {
    return resolveText(c.label, lang);
  }

  flowLegend(c: Category, lang: Lang = getLang()): string {
    return resolveText(c.flowLegend ?? c.label, lang);
  }

  /** 登録されたカテゴリで使われている模様 */
  patterns(): CategoryPattern[] {
    return [...new Set([...this.byId.values()].map((c) => c.pattern).filter((p): p is CategoryPattern => !!p))];
  }
}

/** 模様の id (SVG の <pattern id>) */
export function patternId(p: CategoryPattern): string {
  return `pat-${p}`;
}

/** 使われている模様の <defs>。模様が無ければ空文字列 */
export function renderPatternDefs(patterns: CategoryPattern[]): string {
  if (patterns.length === 0) return "";
  const line = (x1: number, y1: number, x2: number, y2: number) =>
    `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#000" stroke-opacity="0.4" stroke-width="1.2"/>`;
  const defs = patterns.map((p) => {
    const head = `<pattern id="${patternId(p)}" width="6" height="6" patternUnits="userSpaceOnUse"`;
    if (p === "hatch") return `${head} patternTransform="rotate(45)">${line(0, 0, 0, 6)}</pattern>`;
    if (p === "cross") return `${head} patternTransform="rotate(45)">${line(0, 0, 0, 6)}${line(0, 0, 6, 0)}</pattern>`;
    return `${head}><circle cx="3" cy="3" r="1.1" fill="#000" fill-opacity="0.45"/></pattern>`;
  });
  return `<defs>${defs.join("")}</defs>`;
}

/** 矩形に模様を重ねる (模様が無ければ空文字列) */
export function patternOverlay(c: Category, x: number | string, y: number | string, w: number | string, h: number | string): string {
  if (!c.pattern) return "";
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="url(#${patternId(c.pattern)})"/>`;
}
