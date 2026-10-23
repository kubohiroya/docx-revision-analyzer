/**
 * pageLayout.ts
 *
 * 時刻 t の文書 (t 以前の変更を反映し、t より後の変更を元に戻した状態) を、模式的なページに割り付ける。
 * docx-revision-flow の各列のページはこの結果を描く。
 *
 * 再現するもの:
 *  - セクションごとの用紙・余白・とじしろ・ヘッダーとフッターの高さ・文字数と行数 (docGrid)・セクションの始まり方
 *  - 段組み (段数・段の幅・段間・境界線、段区切り、連続のセクション区切りでの段数の切り替えと段の高さの均等化)
 *  - 段落の書式: 段落前後の間隔・行間・インデント・配置・文字の大きさ
 *  - ページ送りの規則: 段落前で改ページ・次の段落と分離しない・段落を分割しない・改ページ時 1 行残して段落を区切らない
 *  - 表: 列の幅・セルの結合・行の高さ・行の途中で改ページしない・見出し行の繰り返し・行の挿入と削除
 *  - 図: 幅と高さ、浮動配置の図の位置と文字の回り込み
 *  - 脚注 (ページの下に確保する高さ) と文末脚注
 *  - 縦書きのセクション (行を右から左へ並べる)
 *  - 最終状態では、Word が保存したときに改ページした位置 (w:lastRenderedPageBreak) を手がかりにする
 *  - 過去の時点の版では、その時点の段落の書式・セクションの設定 (w:pPrChange / w:sectPrChange) を使う
 *
 * 近似にとどまるもの (フォントの寸法や Word の組版の規則そのものが必要なため):
 *  - 1行に入る文字数 (全角は文字の大きさ、半角はその半分として数える。禁則処理・字詰め・単語の折り返しの詳細は扱わない)
 *  - 1行の高さ (よく使われるフォントの概算値を使う)
 *  - 文字の回り込みは図の矩形で行う (外周・内部の輪郭は扱わない)
 *
 * 座標の単位は twip。ページの左上を原点とする。
 */

import type {
  BlockRef,
  DocSection,
  DocxLayoutModel,
  FigureSeg,
  NoteRef,
  Para,
  ParaProps,
  RevRef,
  SectionProps,
  TableSpec,
} from "./docxLayout";

/** ページ内の描画要素 */
export type PageItem =
  | {
      type: "line";
      paraIndex: number;
      x: number;
      y: number;
      w: number;
      h: number;
      /** 行の長さ (行の幅に対する割合、0〜1) */
      fill: number;
      align: "left" | "center" | "right";
      style: Para["kind"];
    }
  | { type: "figure"; paraIndex: number; x: number; y: number; w: number; h: number; float?: boolean }
  /** 表のセルの枠 */
  | { type: "cell"; x: number; y: number; w: number; h: number }
  /** 脚注・文末脚注の行 */
  | { type: "note"; x: number; y: number; w: number; h: number; fill: number }
  /** 繰り返された表の見出し行の行 */
  | { type: "repeat"; x: number; y: number; w: number; h: number; fill: number };

export interface PageLayout {
  /** 用紙の大きさ (twip) */
  width: number;
  height: number;
  /** 縦書きのページ (行の向きが縦) */
  vertical?: boolean;
  items: PageItem[];
  /** 段の境界線・脚注の区切り線 */
  rules?: { x1: number; y1: number; x2: number; y2: number }[];
}

export interface LayoutOptions {
  /** 最終状態 (Word が保存したときの状態) として割り付ける。w:lastRenderedPageBreak を手がかりに使う */
  final?: boolean;
  /** 行の高さの倍率 (総ページ数の補正に使う) */
  lineScale?: number;
}

export interface LayoutResult {
  pages: PageLayout[];
  /** その時点で存在する段落ごとの、見えている文字数 */
  visibleChars: Map<number, number>;
  /** 手がかり (w:lastRenderedPageBreak) で改ページした数 */
  hintBreaks: number;
}

/** 日時の無い挿入は「最初から存在」、日時の無い削除は「最初から削除済み」とみなす */
export function visibleAt(ins: RevRef | undefined, del: RevRef | undefined, t: number): boolean {
  const inserted = !ins || !ins.date || ins.date.getTime() <= t;
  const deleted = !!del && (!del.date || del.date.getTime() <= t);
  return inserted && !deleted;
}

/** 時刻 t の段落の書式 */
export function paraPropsAt(p: Para, t: number): ParaProps {
  return p.propsBefore && p.propsChangedAt && p.propsChangedAt.getTime() > t ? p.propsBefore : p.props;
}

/** 時刻 t のセクションの設定 */
export function sectionAt(s: DocSection, t: number): SectionProps {
  return s.before && s.changedAt && s.changedAt.getTime() > t ? s.before : s.props;
}

/** Word が保存したときの改ページ位置の記録があるか */
export function hasRenderedBreaks(model: DocxLayoutModel): boolean {
  return model.paras.some((p) => p.renderedBreaks?.length);
}

// ---------------------------------------------------------------------------
// 行の折り返し
// ---------------------------------------------------------------------------

/** 折り返しの単位: 全角の文字・半角の単語・空白・タブ */
interface Atom {
  kind: "atom";
  /** 幅 (twip)。タブは 0 (置いた位置で決まる) */
  w: number;
  size: number;
  seg: number;
  space?: boolean;
  tab?: boolean;
}
type Token =
  | Atom
  | { kind: "break"; seg: number; page?: boolean; column?: boolean }
  | { kind: "figure"; seg: number; fig: FigureSeg };

interface TextLine {
  kind: "text";
  /** 行に置いた幅 */
  used: number;
  /** 行の中で最も大きい文字 (pt) */
  size: number;
  /** 行に含まれる断片の範囲 (Para.segs の添字)。空の行は -1 */
  segFrom: number;
  segTo: number;
  /** 行の後で改ページ / 段区切り */
  forced?: "page" | "column";
  /** 改ページ・段区切りだけで、行を占めない */
  empty?: boolean;
}
type Line = TextLine | { kind: "figure"; fig: FigureSeg; seg: number };

interface Metrics {
  /** 全角1文字の送り (twip)。文字の大きさ size (pt) の文字 */
  full(size: number): number;
  half(size: number): number;
  tabStop: number;
}

function tokenize(p: Para, t: number, m: Metrics, defaultSize: number): Token[] {
  const tokens: Token[] = [];
  p.segs.forEach((s, seg) => {
    if (!visibleAt(s.ins, s.del, t)) return;
    if (s.kind === "lineBreak" || s.kind === "pageBreak") {
      tokens.push({ kind: "break", seg, page: s.kind === "pageBreak", column: s.column });
      return;
    }
    if (s.kind === "figure") {
      tokens.push({ kind: "figure", seg, fig: s });
      return;
    }
    if (s.kind !== "text" || s.hidden) return;
    const size = s.sizePt ?? defaultSize;
    let word: Atom | undefined;
    for (const ch of s.text) {
      const cp = ch.codePointAt(0)!;
      if (ch === "\t") {
        word = undefined;
        tokens.push({ kind: "atom", w: 0, size, seg, tab: true });
      } else if (ch === " " || ch === " ") {
        word = undefined;
        tokens.push({ kind: "atom", w: m.half(size), size, seg, space: true });
      } else if (cp >= 0x1100) {
        word = undefined;
        tokens.push({ kind: "atom", w: m.full(size), size, seg });
      } else if (word) {
        word.w += m.half(size);
      } else {
        word = { kind: "atom", w: m.half(size), size, seg };
        tokens.push(word);
      }
    }
  });
  return tokens;
}

/** 段落の行を1行ずつ取り出す。clone() で先読みできる */
class Breaker {
  private i = 0;
  /** いまの単位のうち、前の行に置いた幅 (行より長い単語を分けたとき) */
  private taken = 0;
  constructor(
    private readonly tokens: Token[],
    private readonly tabStop: number
  ) {}

  clone(): Breaker {
    const b = new Breaker(this.tokens, this.tabStop);
    b.i = this.i;
    b.taken = this.taken;
    return b;
  }

  get done(): boolean {
    return this.i >= this.tokens.length;
  }

  next(width: number): Line | undefined {
    if (this.done) return undefined;
    const first = this.tokens[this.i];
    if (first.kind === "figure") {
      this.i++;
      return { kind: "figure", fig: first.fig, seg: first.seg };
    }
    const line: TextLine = { kind: "text", used: 0, size: 0, segFrom: -1, segTo: -1 };
    const mark = (seg: number) => {
      if (line.segFrom < 0) line.segFrom = seg;
      line.segTo = seg;
    };
    while (this.i < this.tokens.length) {
      const tk = this.tokens[this.i];
      if (tk.kind === "break") {
        this.i++;
        if (tk.page || tk.column) {
          line.forced = tk.page ? "page" : "column";
          if (line.segFrom < 0) line.empty = true;
        }
        mark(tk.seg);
        return line;
      }
      if (tk.kind === "figure") return line;
      const w = tk.tab ? this.tabStop - (line.used % this.tabStop) : tk.w - this.taken;
      if (line.used + w > width && line.used > 0 && !tk.space) return line;
      if (line.used + w > width && line.used === 0 && !tk.tab) {
        // 行より長い単語は、行の幅で分ける
        this.taken += width;
        line.used = width;
        line.size = Math.max(line.size, tk.size);
        mark(tk.seg);
        return line;
      }
      line.used += w;
      line.size = Math.max(line.size, tk.size);
      mark(tk.seg);
      this.i++;
      this.taken = 0;
    }
    return line;
  }
}

// ---------------------------------------------------------------------------
// 割り付け
// ---------------------------------------------------------------------------

interface Rect {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

/** 段組みの均等化で動かす単位 (行・図・表の行) */
interface Unit {
  items: PageItem[];
  col: number;
  /** 単位の上端 (間隔の後) と高さ、直前の間隔 */
  y: number;
  h: number;
  gap: number;
}

/** 測った段落・表 (ページ送りをしない、表のセルの中など) */
interface Measured {
  items: PageItem[];
  h: number;
}

interface MeasuredCell {
  x: number;
  w: number;
  gridCol: number;
  merge?: "restart" | "continue";
  items: PageItem[];
}

interface MeasuredRow {
  h: number;
  header: boolean;
  cantSplit: boolean;
  cells: MeasuredCell[];
}

function shift(items: PageItem[], dx: number, dy: number): void {
  for (const it of items) {
    it.x += dx;
    it.y += dy;
  }
}

class Layouter {
  readonly pages: PageLayout[] = [];
  readonly visibleChars = new Map<number, number>();
  hintBreaks = 0;

  private page!: PageLayout;
  private sp!: SectionProps;
  /** 論理座標 (縦書きでは行の方向を横にした座標) でのページの大きさと本文の範囲 */
  private lw = 0;
  private lh = 0;
  private top = 0;
  private bottom = 0;
  private left = 0;
  private right = 0;
  private cols: { x: number; w: number }[] = [];
  private col = 0;
  /** いまの段組みの範囲の上端 (連続のセクション区切りの後は、前の段組みの下から) */
  private regionTop = 0;
  private colStartY = 0;
  private y = 0;
  /** 段ごとの、脚注を除いた本文の下端 */
  private colBottom: number[] = [];
  /** 段ごとの脚注の行 (高さと長さ) */
  private notes: { h: number; fill: number }[][] = [];
  private exclusions: Rect[] = [];
  private units: Unit[] = [];
  /** 直前の段落 (同じスタイルの段落の間隔を詰めるため) と、その後の間隔 */
  private prev?: { styleId?: string; contextual: boolean; after: number };
  /** 本文で参照された文末脚注 (文書の最後に置く) */
  private endnotes: NoteRef[] = [];

  constructor(
    private readonly model: DocxLayoutModel,
    private readonly t: number,
    private readonly opts: LayoutOptions
  ) {}

  // ---- 文字の大きさと行の高さ ------------------------------------------------

  private gridLines(pp: ParaProps): boolean {
    return this.sp.gridType !== "default" && pp.snapToGrid;
  }

  private gridChars(pp: ParaProps): boolean {
    return (this.sp.gridType === "linesAndChars" || this.sp.gridType === "snapToChars") && pp.snapToGrid;
  }

  /** 文字数を指定したときの全角1文字の送り */
  private charPitch(): number {
    return this.model.baseSizePt * 20 + (this.sp.charSpace * 20) / 4096;
  }

  private metrics(pp: ParaProps): Metrics {
    const grid = this.gridChars(pp);
    const pitch = this.charPitch();
    return {
      full: (size) => {
        const w = size * 20;
        return grid && pitch > 0 ? Math.max(1, Math.ceil(w / pitch - 0.05)) * pitch : w;
      },
      half: (size) => size * 10,
      tabStop: this.model.defaultTabStop,
    };
  }

  /** 字数で指定したインデントの1字 */
  private charUnit(pp: ParaProps): number {
    return this.gridChars(pp) ? this.charPitch() : pp.sizePt * 20;
  }

  private lineHeight(size: number, pp: ParaProps): number {
    const natural = size * 20 * pp.lineFactor;
    const pitch = this.sp.linePitch;
    const single = this.gridLines(pp) && pitch > 0 ? pitch * Math.max(1, Math.ceil(natural / pitch - 0.1)) : natural;
    const scale = this.opts.lineScale ?? 1;
    if (pp.lineRule === "exact") return pp.line;
    if (pp.lineRule === "atLeast") return Math.max(pp.line, single * scale);
    return ((single * pp.line) / 240) * scale;
  }

  private spacing(pp: ParaProps, which: "before" | "after"): number {
    const lines = which === "before" ? pp.spaceBeforeLines : pp.spaceAfterLines;
    if (lines !== undefined) return (lines / 100) * (this.sp.gridType !== "default" ? this.sp.linePitch : 240);
    return which === "before" ? pp.spaceBefore : pp.spaceAfter;
  }

  private indents(pp: ParaProps): { left: number; right: number; first: number } {
    const unit = this.charUnit(pp);
    return {
      left: pp.indLeftChars !== undefined ? (pp.indLeftChars / 100) * unit : pp.indLeft,
      right: pp.indRightChars !== undefined ? (pp.indRightChars / 100) * unit : pp.indRight,
      first: pp.indFirstChars !== undefined ? (pp.indFirstChars / 100) * unit : pp.indFirst,
    };
  }

  // ---- ページと段 ------------------------------------------------------------

  private get colHeight(): number {
    return this.bottom - this.regionTop;
  }

  private get atTop(): boolean {
    return this.y <= this.colStartY + 0.5;
  }

  /** ページの先頭 (まだ何も置いていない) */
  private get pageEmpty(): boolean {
    return this.page.items.length === 0 && this.col === 0 && this.atTop;
  }

  private setFrame(sp: SectionProps): void {
    this.sp = sp;
    if (sp.vertical) {
      // 縦書き: 行の方向 (上→下) を横に、行を送る方向 (右→左) を下にした座標で組む
      this.lw = sp.pageHeight;
      this.lh = sp.pageWidth;
      this.top = sp.marginRight;
      this.bottom = sp.pageWidth - sp.marginLeft;
      this.left = sp.marginTop;
      this.right = sp.pageHeight - sp.marginBottom;
    } else {
      this.lw = sp.pageWidth;
      this.lh = sp.pageHeight;
      // ヘッダー・フッターが余白に収まらなければ本文を押し下げる
      this.top = Math.max(sp.marginTop, sp.headerHeight > 0 ? sp.headerDistance + sp.headerHeight : 0);
      this.bottom = sp.pageHeight - Math.max(sp.marginBottom, sp.footerHeight > 0 ? sp.footerDistance + sp.footerHeight : 0);
      this.left = sp.marginLeft + sp.gutter;
      this.right = sp.pageWidth - sp.marginRight;
    }
    const avail = Math.max(1, this.right - this.left);
    const { num, space, widths } = sp.columns;
    this.cols = [];
    if (widths && widths.length === num) {
      let x = this.left;
      const total = widths.reduce((n, c, i) => n + c.width + (i < num - 1 ? c.space : 0), 0);
      const k = total > avail ? avail / total : 1;
      for (const c of widths) {
        this.cols.push({ x, w: c.width * k });
        x += (c.width + c.space) * k;
      }
    } else {
      const w = Math.max(1, (avail - space * (num - 1)) / num);
      for (let i = 0; i < num; i++) this.cols.push({ x: this.left + i * (w + space), w });
    }
    this.colBottom = this.cols.map(() => this.bottom);
    this.notes = this.cols.map(() => []);
  }

  private newPage(): void {
    if (this.page) this.finishPage();
    this.page = { width: this.sp.pageWidth, height: this.sp.pageHeight, items: [] };
    if (this.sp.vertical) this.page.vertical = true;
    this.pages.push(this.page);
    this.regionTop = this.top;
    this.col = 0;
    this.y = this.colStartY = this.top;
    this.colBottom = this.cols.map(() => this.bottom);
    this.notes = this.cols.map(() => []);
    this.exclusions = [];
    this.units = [];
  }

  private advanceColumn(): void {
    if (this.col < this.cols.length - 1) {
      this.col++;
      this.y = this.colStartY = this.regionTop;
    } else {
      this.newPage();
    }
  }

  /** 段組みの範囲を閉じる: 段の境界線を引く */
  private closeRegion(bottom: number): void {
    if (!this.sp.columns.separator || this.cols.length < 2 || bottom <= this.regionTop) return;
    for (let i = 0; i < this.cols.length - 1; i++) {
      const x = (this.cols[i].x + this.cols[i].w + this.cols[i + 1].x) / 2;
      (this.page.rules ??= []).push({ x1: x, y1: this.regionTop, x2: x, y2: bottom });
    }
  }

  private usedBottom(): number {
    return this.units.reduce((b, u) => Math.max(b, u.y + u.h), this.regionTop);
  }

  private finishPage(): void {
    const page = this.page;
    this.closeRegion(this.cols.length > 1 ? this.bottom : this.usedBottom());
    // 脚注: 段の下から積む
    this.notes.forEach((lines, c) => {
      if (lines.length === 0) return;
      const total = lines.reduce((n, l) => n + l.h, 0);
      const col = this.cols[c];
      let y = this.bottom - total;
      (page.rules ??= []).push({ x1: col.x, y1: y - 60, x2: col.x + col.w * 0.33, y2: y - 60 });
      for (const l of lines) {
        page.items.push({ type: "note", x: col.x, y, w: col.w, h: l.h, fill: l.fill });
        y += l.h;
      }
    });
    if (page.vertical) {
      // 論理座標から縦書きのページの座標へ: 行の方向は上→下、行は右から左へ
      const W = page.width;
      for (const it of page.items) {
        const { x, y, w, h } = it;
        Object.assign(it, { x: W - (y + h), y: x, w: h, h: w });
      }
      for (const r of page.rules ?? []) Object.assign(r, { x1: W - r.y1, y1: r.x1, x2: W - r.y2, y2: r.x2 });
    }
  }

  private startSection(index: number, first: boolean): void {
    const sp = sectionAt(this.model.sections[index], this.t);
    if (first) {
      this.setFrame(sp);
      this.newPage();
      return;
    }
    const prev = this.sp;
    const samePage =
      prev.pageWidth === sp.pageWidth && prev.pageHeight === sp.pageHeight && prev.vertical === sp.vertical;
    if (sp.start === "continuous" && samePage) {
      // 連続: 前の段組みの段の高さをそろえ、その下から新しい段組みで組む
      const bottom = this.cols.length > 1 ? this.balance() : this.y;
      this.closeRegion(bottom);
      const keepTop = this.top;
      const keepBottom = this.bottom;
      // このページの脚注は引き継ぐ (新しい段組みの1段目に置く)
      const notes = this.notes.flat();
      const reserved = Math.max(...this.colBottom.map((b) => keepBottom - b));
      this.setFrame(sp);
      this.top = keepTop;
      this.bottom = keepBottom;
      this.notes[0] = notes;
      this.colBottom = this.cols.map(() => this.bottom - reserved);
      this.regionTop = bottom;
      this.col = 0;
      this.y = this.colStartY = bottom;
      this.units = [];
      return;
    }
    if (sp.start === "nextColumn" && samePage && prev.columns.num === sp.columns.num && this.col < this.cols.length - 1) {
      this.setFrame(sp);
      this.advanceColumn();
      return;
    }
    if (this.page.items.length === 0 && !this.page.rules) {
      // まだ何も置いていないページは、新しいセクションの最初のページとして使う
      this.pages.pop();
      this.page = undefined as unknown as PageLayout;
      this.setFrame(sp);
      this.newPage();
    } else {
      this.finishPage();
      this.setFrame(sp);
      this.page = undefined as unknown as PageLayout;
      this.newPage();
    }
    // 奇数・偶数ページから始める: ページ番号が合わなければ白紙のページを1枚入れる
    const n = this.pages.length;
    if ((sp.start === "oddPage" && n % 2 === 0) || (sp.start === "evenPage" && n % 2 === 1)) this.newPage();
  }

  /**
   * いまの段組みの範囲 (このページの、このセクションの部分) の段の高さをそろえる。
   * 置いた単位を、同じ順序のまま、できるだけ低い段の高さで詰め直す。範囲の下端を返す
   */
  private balance(): number {
    const units = this.units;
    if (units.length === 0) return this.regionTop;
    const n = this.cols.length;
    const assign = (H: number): { col: number; y: number }[] | undefined => {
      const out: { col: number; y: number }[] = [];
      let col = 0;
      let y = this.regionTop;
      for (const u of units) {
        let gap = y === this.regionTop ? 0 : u.gap;
        if (y + gap + u.h > this.regionTop + H + 0.5 && y > this.regionTop) {
          col++;
          y = this.regionTop;
          gap = 0;
          if (col >= n) return undefined;
        }
        out.push({ col, y: y + gap });
        y += gap + u.h;
      }
      return out;
    };
    let lo = Math.max(...units.map((u) => u.h));
    let hi = units.reduce((s, u) => s + u.gap + u.h, 0);
    if (!assign(hi)) return this.usedBottom();
    for (let k = 0; k < 20 && hi - lo > 1; k++) {
      const mid = (lo + hi) / 2;
      if (assign(mid)) hi = mid;
      else lo = mid;
    }
    const placement = assign(hi)!;
    let bottom = this.regionTop;
    units.forEach((u, i) => {
      const p = placement[i];
      shift(u.items, this.cols[p.col].x - this.cols[u.col].x, p.y - u.y);
      u.col = p.col;
      u.y = p.y;
      bottom = Math.max(bottom, p.y + u.h);
    });
    return bottom;
  }

  /** 単位を置く (items の座標はすでに決まっているもの) */
  private addUnit(items: PageItem[], h: number, gap: number): void {
    this.page.items.push(...items);
    this.units.push({ items, col: this.col, y: this.y, h, gap });
  }

  // ---- 文字の回り込み ----------------------------------------------------------

  /** 高さ h の行を y に置くときに使える横の範囲。図を避けられなければ y を下げる */
  private span(y: number, h: number, x: number, w: number): { x: number; w: number; y: number } {
    for (let guard = 0; guard < 50; guard++) {
      const hit = this.exclusions.filter((e) => e.y1 < y + h && e.y2 > y && e.x1 < x + w && e.x2 > x);
      if (hit.length === 0) return { x, w, y };
      // 図の左右に残る区間のうち、最も広いもの
      const cuts = hit.map((e) => [Math.max(x, e.x1), Math.min(x + w, e.x2)]).sort((a, b) => a[0] - b[0]);
      let best = { x, w: 0 };
      let cur = x;
      for (const [a, b] of cuts) {
        if (a - cur > best.w) best = { x: cur, w: a - cur };
        cur = Math.max(cur, b);
      }
      if (x + w - cur > best.w) best = { x: cur, w: x + w - cur };
      if (best.w >= Math.max(w * 0.25, 720)) return { ...best, y };
      y = Math.min(...hit.map((e) => e.y2));
    }
    return { x, w, y };
  }

  private placeFloat(fig: FigureSeg, paraIndex: number, paraTop: number): void {
    const f = fig.float!;
    const col = this.cols[this.col];
    const w = Math.min(fig.widthTwips ?? col.w * 0.7, this.lw);
    const h = Math.min(fig.heightTwips, this.lh);
    const rangeH: Record<string, [number, number]> = {
      page: [0, this.lw],
      margin: [this.left, this.right],
      leftMargin: [0, this.left],
      rightMargin: [this.right, this.lw],
      insideMargin: [0, this.left],
      outsideMargin: [this.right, this.lw],
    };
    const rangeV: Record<string, [number, number]> = {
      page: [0, this.lh],
      margin: [this.top, this.bottom],
      topMargin: [0, this.top],
      bottomMargin: [this.bottom, this.lh],
    };
    const [a, b] = rangeH[f.relH] ?? [col.x, col.x + col.w];
    const [c, d] = rangeV[f.relV] ?? [paraTop, paraTop + h];
    const align = (lo: number, hi: number, size: number, v: string | undefined) =>
      v === "center" ? lo + (hi - lo - size) / 2 : v === "right" || v === "bottom" || v === "outside" ? hi - size : lo;
    let x = f.alignH ? align(a, b, w, f.alignH) : a + (f.offH ?? 0);
    let y = f.alignV ? align(c, d, h, f.alignV) : c + (f.offV ?? 0);
    x = Math.min(Math.max(0, x), this.lw - w);
    y = Math.min(Math.max(0, y), this.lh - h);
    this.page.items.push({ type: "figure", paraIndex, x, y, w, h, float: true });
    if (f.wrap === "topBottom") {
      this.exclusions.push({ x1: -Infinity, y1: y - f.dist.t, x2: Infinity, y2: y + h + f.dist.b });
    } else if (f.wrap === "square") {
      this.exclusions.push({ x1: x - f.dist.l, y1: y - f.dist.t, x2: x + w + f.dist.r, y2: y + h + f.dist.b });
    }
  }

  // ---- 脚注 ------------------------------------------------------------------

  private noteLines(ref: NoteRef, width: number): { h: number; fill: number }[] {
    const info = (ref.note === "footnote" ? this.model.footnotes : this.model.endnotes).get(ref.id);
    const size = this.model.noteSizePt;
    const perLine = Math.max(1, width / (size * 20));
    const lineH = size * 20 * 1.2 * (this.opts.lineScale ?? 1);
    const total = info ? info.width : 0;
    const count = Math.max(info?.paras ?? 1, Math.ceil(total / perLine));
    const lines: { h: number; fill: number }[] = [];
    let rest = total;
    for (let i = 0; i < count; i++) {
      lines.push({ h: lineH, fill: Math.min(1, Math.max(0.1, rest / perLine)) });
      rest -= perLine;
    }
    return lines;
  }

  // ---- 段落 ------------------------------------------------------------------

  private exists(p: Para): boolean {
    const visible = p.segs.some((s) => visibleAt(s.ins, s.del, this.t));
    return visible || visibleAt(p.markIns, p.markDel, this.t);
  }

  private countChars(p: Para): void {
    let n = 0;
    for (const s of p.segs) if (s.kind === "text" && visibleAt(s.ins, s.del, this.t)) n += s.chars;
    this.visibleChars.set(p.index, n);
  }

  private styleOf(p: Para, pp: ParaProps): Para["kind"] {
    return p.tableId !== undefined ? "table" : pp.headingLevel !== undefined ? "heading" : "body";
  }

  private alignOf(pp: ParaProps): "left" | "center" | "right" {
    return pp.jc === "center" ? "center" : pp.jc === "right" ? "right" : "left";
  }

  private breaker(p: Para, pp: ParaProps): Breaker {
    return new Breaker(tokenize(p, this.t, this.metrics(pp), pp.sizePt), this.model.defaultTabStop);
  }

  private lineH(line: Line, pp: ParaProps): number {
    return line.kind === "figure" ? Math.min(line.fig.heightTwips, this.lh) : this.lineHeight(line.size || pp.sizePt, pp);
  }

  /** 段落の行の高さ (先読み。幅 width で折り返したもの) */
  private nominalLines(p: Para, pp: ParaProps, width: number): number[] {
    const ind = this.indents(pp);
    const b = this.breaker(p, pp);
    const out: number[] = [];
    for (let line = b.next(width - ind.left - ind.right - ind.first); line; line = b.next(width - ind.left - ind.right)) {
      if (line.kind === "text" && line.empty) continue;
      if (line.kind === "figure" && line.fig.float) continue;
      out.push(this.lineH(line, pp));
    }
    if (out.length === 0) out.push(this.lineHeight(pp.sizePt, pp));
    return out;
  }

  /** 段落の前の間隔 (同じスタイルの段落の間を詰める設定と、段の先頭での省略を反映) */
  private gapBefore(pp: ParaProps): number {
    let before = this.spacing(pp, "before");
    const prev = this.prev;
    if (prev && prev.styleId === pp.styleId) {
      if (pp.contextualSpacing) before = 0;
      if (prev.contextual && !this.atTop) {
        this.y -= prev.after;
        prev.after = 0;
      }
    }
    return this.atTop && this.pages.length + this.col + this.units.length > 1 ? 0 : before;
  }

  /** 本文の段落を割り付ける */
  private placePara(p: Para): void {
    if (!this.exists(p)) return;
    this.countChars(p);
    const pp = paraPropsAt(p, this.t);
    if (pp.pageBreakBefore && !this.pageEmpty) this.newPage();
    const hints = this.opts.final ? p.renderedBreaks ?? [] : [];
    const firstVisible = p.segs.findIndex((s) => visibleAt(s.ins, s.del, this.t));
    if (hints.some((h) => h <= Math.max(0, firstVisible))) this.hintBreak();

    const ind = this.indents(pp);
    const nominal = this.nominalLines(p, pp, this.cols[this.col].w);
    let gap = this.gapBefore(pp);
    // 段落を分割しない・1行残して区切らない
    let splitAfter = -1;
    if (!this.atTop) {
      const avail = this.colBottom[this.col] - this.y - gap;
      const total = nominal.reduce((a, b) => a + b, 0);
      let fit = 0;
      for (let acc = 0; fit < nominal.length && acc + nominal[fit] <= avail; fit++) acc += nominal[fit];
      if (fit < nominal.length) {
        if (pp.keepLines && total + gap <= this.colHeight) {
          this.advanceColumn();
          gap = this.gapBefore(pp);
        } else if (pp.widowControl && nominal.length > 1) {
          if (fit === 1) {
            this.advanceColumn();
            gap = this.gapBefore(pp);
          } else if (fit >= 2 && fit === nominal.length - 1) {
            splitAfter = fit - 1;
          }
        }
      }
    }
    this.y += gap;
    const paraTop = this.y;
    for (const s of p.segs) {
      if (s.kind === "figure" && s.float && visibleAt(s.ins, s.del, this.t)) this.placeFloat(s, p.index, paraTop);
    }

    const b = this.breaker(p, pp);
    const style = this.styleOf(p, pp);
    const align = this.alignOf(pp);
    let pending = (p.noteRefs ?? []).filter((r) => visibleAt(r.ins, r.del, this.t));
    let lineNo = 0;
    for (;;) {
      const col = this.cols[this.col];
      const est = this.lineHeight(pp.sizePt, pp);
      const first = lineNo === 0;
      const sp0 = this.span(this.y, est, col.x + ind.left, col.w - ind.left - ind.right);
      const line = b.next(sp0.w - (first ? ind.first : 0));
      if (!line) break;
      if (line.kind === "figure" && line.fig.float) continue;
      if (line.kind === "text" && line.empty) {
        if (line.forced === "page") this.newPage();
        else if (line.forced === "column") this.advanceColumn();
        continue;
      }
      const h = this.lineH(line, pp);
      // 手がかり: この行の中で Word が改ページしていた
      if (line.kind === "text" && hints.some((hk) => hk > 0 && hk >= line.segFrom && hk <= line.segTo)) this.hintBreak();
      if (splitAfter >= 0 && lineNo === splitAfter && !this.atTop) this.advanceColumn();
      // この行までに参照された脚注 (参照の直前の断片がこの行にあるもの)
      const lineRefs = line.kind === "text" ? pending.filter((r) => r.segIndex - 1 <= line.segTo) : [];
      pending = pending.filter((r) => !lineRefs.includes(r));
      for (const r of lineRefs) if (r.note === "endnote") this.endnotes.push(r);
      const noteLines = lineRefs.filter((r) => r.note === "footnote").flatMap((r) => this.noteLines(r, col.w));
      const noteH = noteLines.reduce((n, l) => n + l.h, 0) + (noteLines.length && !this.notes[this.col].length ? 240 : 0);
      if (this.y + h + noteH > this.colBottom[this.col] + 0.5 && !this.atTop) this.advanceColumn();
      const c = this.cols[this.col];
      let sp = this.span(this.y, h, c.x + ind.left, c.w - ind.left - ind.right);
      if (sp.y + h > this.colBottom[this.col] + 0.5 && sp.y > this.y && !this.atTop) {
        this.advanceColumn();
        const c2 = this.cols[this.col];
        sp = this.span(this.y, h, c2.x + ind.left, c2.w - ind.left - ind.right);
      }
      this.y = sp.y;
      let item: PageItem;
      if (line.kind === "figure") {
        const w = Math.min(line.fig.widthTwips ?? sp.w * 0.7, sp.w);
        const x = align === "center" ? sp.x + (sp.w - w) / 2 : align === "right" ? sp.x + sp.w - w : sp.x;
        item = { type: "figure", paraIndex: p.index, x, y: this.y, w, h };
      } else {
        const x = sp.x + (first ? ind.first : 0);
        const w = Math.max(1, sp.w - (first ? ind.first : 0));
        item = { type: "line", paraIndex: p.index, x, y: this.y, w, h, fill: Math.min(1, line.used / w), align, style };
      }
      this.addUnit([item], h, lineNo === 0 ? gap : 0);
      if (noteLines.length) {
        if (!this.notes[this.col].length) this.colBottom[this.col] -= 240;
        this.notes[this.col].push(...noteLines);
        this.colBottom[this.col] -= noteLines.reduce((n, l) => n + l.h, 0);
      }
      this.y += h;
      lineNo++;
      if (line.kind === "text" && line.forced === "page") this.newPage();
      else if (line.kind === "text" && line.forced === "column") this.advanceColumn();
    }
    // 行に割り当てられなかった参照 (段落の末尾など)
    for (const r of pending) if (r.note === "endnote") this.endnotes.push(r);
    if (lineNo === 0) {
      // 空の段落も1行分の高さを占める
      const h = this.lineHeight(pp.sizePt, pp);
      if (this.y + h > this.colBottom[this.col] + 0.5 && !this.atTop) this.advanceColumn();
      const c = this.cols[this.col];
      const item: PageItem = {
        type: "line",
        paraIndex: p.index,
        x: c.x + ind.left,
        y: this.y,
        w: Math.max(1, c.w - ind.left - ind.right),
        h,
        fill: 0,
        align,
        style,
      };
      this.addUnit([item], h, gap);
      this.y += h;
    }
    const after = this.spacing(pp, "after");
    this.y += after;
    this.prev = { styleId: pp.styleId, contextual: pp.contextualSpacing, after };
  }

  /** 手がかりの位置で、段の大半が埋まっていれば次の段 (ページ) へ送る */
  private hintBreak(): void {
    if (this.atTop) return;
    const filled = (this.y - this.colStartY) / Math.max(1, this.colBottom[this.col] - this.colStartY);
    if (filled >= 0.5) {
      this.advanceColumn();
      this.hintBreaks++;
    }
  }

  /** 次の段落と分離しない段落が続くとき、その全体と次の段落の1行目が段に入らなければ先に送る */
  private keepWithNext(blocks: BlockRef[], i: number): number {
    const paras = this.model.paras;
    let total = 0;
    let j = i;
    const width = this.cols[this.col].w;
    for (; j < blocks.length; j++) {
      const b = blocks[j];
      if (b.kind !== "para") break;
      const p = paras[b.index];
      if (!this.exists(p)) continue;
      const pp = paraPropsAt(p, this.t);
      const lines = this.nominalLines(p, pp, width);
      if (!pp.keepNext) {
        total += this.spacing(pp, "before") + lines[0];
        break;
      }
      total += this.spacing(pp, "before") + lines.reduce((a, c) => a + c, 0) + this.spacing(pp, "after");
    }
    if (j === blocks.length || blocks[j]?.kind === "table") total += this.lineHeight(this.model.baseSizePt, paras[(blocks[i] as { index: number }).index].props);
    if (!this.atTop && this.y + total > this.colBottom[this.col] && total <= this.colHeight) this.advanceColumn();
    return j;
  }

  // ---- 表 --------------------------------------------------------------------

  /** ページ送りをせずに段落と表を縦に積む (表のセルの中) */
  private stack(blocks: BlockRef[], x: number, width: number): Measured {
    const items: PageItem[] = [];
    let y = 0;
    let prev: { styleId?: string; contextual: boolean; after: number } | undefined;
    for (const b of blocks) {
      if (b.kind === "table") {
        const rows = this.measureTable(this.model.tables[b.table], x, width);
        const merged = new Map<number, PageItem & { type: "cell" }>();
        for (const r of rows) {
          this.emitRow(r, y, items, merged);
          y += r.h;
        }
        prev = undefined;
        continue;
      }
      const p = this.model.paras[b.index];
      if (!this.exists(p)) continue;
      this.countChars(p);
      const pp = paraPropsAt(p, this.t);
      const ind = this.indents(pp);
      let before = this.spacing(pp, "before");
      if (prev && prev.styleId === pp.styleId) {
        if (pp.contextualSpacing) before = 0;
        if (prev.contextual) y -= prev.after;
      }
      y += items.length ? before : 0;
      const br = this.breaker(p, pp);
      const style = this.styleOf(p, pp);
      const align = this.alignOf(pp);
      let lineNo = 0;
      const lw = Math.max(1, width - ind.left - ind.right);
      for (let line = br.next(lw - ind.first); line; line = br.next(lw)) {
        if (line.kind === "text" && line.empty) continue;
        const h = this.lineH(line, pp);
        if (line.kind === "figure") {
          const w = Math.min(line.fig.widthTwips ?? lw * 0.7, lw);
          const fx = align === "center" ? x + ind.left + (lw - w) / 2 : align === "right" ? x + ind.left + lw - w : x + ind.left;
          items.push({ type: "figure", paraIndex: p.index, x: fx, y, w, h });
        } else {
          const first = lineNo === 0 ? ind.first : 0;
          const w = Math.max(1, lw - first);
          items.push({ type: "line", paraIndex: p.index, x: x + ind.left + first, y, w, h, fill: Math.min(1, line.used / w), align, style });
        }
        y += h;
        lineNo++;
      }
      if (lineNo === 0) {
        const h = this.lineHeight(pp.sizePt, pp);
        items.push({ type: "line", paraIndex: p.index, x: x + ind.left, y, w: lw, h, fill: 0, align, style });
        y += h;
      }
      const after = this.spacing(pp, "after");
      y += after;
      prev = { styleId: pp.styleId, contextual: pp.contextualSpacing, after };
    }
    // 最後の段落の後の間隔はセルの高さに含めない
    if (prev) y -= prev.after;
    return { items, h: Math.max(0, y) };
  }

  /** 表の行を測る (列の幅・セルの中身・行の高さ・縦の結合) */
  private measureTable(spec: TableSpec, x0: number, avail: number): MeasuredRow[] {
    const rows = spec.rows.filter((r) => visibleAt(r.ins, r.del, this.t));
    let grid = spec.grid.filter((g) => g > 0);
    const cellsPerRow = Math.max(1, ...rows.map((r) => r.gridBefore + r.cells.reduce((n, c) => n + c.gridSpan, 0)));
    if (grid.length < cellsPerRow) {
      const fromCells = rows[0]?.cells.every((c) => c.width) ? rows[0].cells.flatMap((c) => Array(c.gridSpan).fill(c.width! / c.gridSpan)) : [];
      grid = fromCells.length >= cellsPerRow ? fromCells : Array(cellsPerRow).fill(avail / cellsPerRow);
    }
    let total = grid.reduce((a, b) => a + b, 0);
    // 本文の幅を超える表は、幅に収める
    if (total > avail * 1.02) {
      grid = grid.map((g) => (g * avail) / total);
      total = avail;
    }
    const tableX =
      spec.jc === "center" ? x0 + (avail - total) / 2 : spec.jc === "right" ? x0 + avail - total : x0 + Math.min(spec.indent, Math.max(0, avail - total));
    const colX = (i: number) => tableX + grid.slice(0, i).reduce((a, b) => a + b, 0);
    const minLine = this.model.baseSizePt * 20 * 1.2;

    const measured: MeasuredRow[] = [];
    /** 縦に結合したセルの、先頭の行と必要な高さ (列ごと) */
    const spans: { row: number; gridCol: number; need: number }[] = [];
    rows.forEach((row, ri) => {
      let gc = row.gridBefore;
      let h = 0;
      const cells: MeasuredCell[] = [];
      for (const cell of row.cells) {
        const x = colX(gc);
        const w = grid.slice(gc, gc + cell.gridSpan).reduce((a, b) => a + b, 0);
        const m = { ...spec.cellMargin, ...Object.fromEntries(Object.entries(cell.margin ?? {}).filter(([, v]) => v !== undefined)) };
        const mc: MeasuredCell = { x, w, gridCol: gc, merge: cell.vMerge, items: [] };
        if (cell.vMerge !== "continue") {
          const content = this.stack(cell.blocks, x + m.left, Math.max(1, w - m.left - m.right));
          shift(content.items, 0, m.top);
          mc.items = content.items;
          const need = content.h + m.top + m.bottom;
          if (cell.vMerge === "restart") spans.push({ row: ri, gridCol: gc, need });
          else h = Math.max(h, need);
        }
        cells.push(mc);
        gc += cell.gridSpan;
      }
      if (h === 0) h = minLine;
      if (row.height !== undefined && row.heightRule === "exact") h = row.height;
      else if (row.height !== undefined && row.heightRule === "atLeast") h = Math.max(h, row.height);
      measured.push({ h, header: row.header, cantSplit: row.cantSplit, cells });
    });
    // 縦に結合したセルの中身が、結合した行の合計より高ければ、最後の行を高くする
    for (const s of spans) {
      let end = s.row;
      while (end + 1 < measured.length && measured[end + 1].cells.some((c) => c.gridCol === s.gridCol && c.merge === "continue")) end++;
      const sum = measured.slice(s.row, end + 1).reduce((a, r) => a + r.h, 0);
      if (s.need > sum) measured[end].h += s.need - sum;
    }
    return measured;
  }

  /** 行の中身とセルの枠を items に出す。縦に結合したセルの枠は、merged で前の行の枠を伸ばす */
  private emitRow(r: MeasuredRow, y: number, items: PageItem[], merged: Map<number, PageItem & { type: "cell" }>, repeat = false): void {
    for (const c of r.cells) {
      const open = merged.get(c.gridCol);
      if (c.merge === "continue" && open) {
        open.h += r.h;
      } else {
        const box: PageItem & { type: "cell" } = { type: "cell", x: c.x, y, w: c.w, h: r.h };
        items.push(box);
        if (c.merge === "restart") merged.set(c.gridCol, box);
        else merged.delete(c.gridCol);
      }
      for (const it of c.items) {
        if (repeat && it.type === "line") items.push({ type: "repeat", x: it.x, y: it.y + y, w: it.w, h: it.h, fill: it.fill });
        else if (!repeat) items.push({ ...it, y: it.y + y });
      }
    }
  }

  /** 本文の表を割り付ける (行ごとにページ送りをする) */
  private placeTable(spec: TableSpec): void {
    const col = this.cols[this.col];
    const rows = this.measureTable(spec, col.x, col.w);
    if (rows.length === 0) return;
    const headers = rows.filter((r, i) => r.header && rows.slice(0, i).every((x) => x.header));
    let merged = new Map<number, PageItem & { type: "cell" }>();
    const dx = () => this.cols[this.col].x - col.x;
    const placeHeaders = () => {
      merged = new Map();
      for (const r of headers) {
        const items: PageItem[] = [];
        this.emitRow(r, this.y, items, merged, true);
        shift(items, dx(), 0);
        this.addUnit(items, r.h, 0);
        this.y += r.h;
      }
    };
    const next = () => {
      this.advanceColumn();
      placeHeaders();
    };
    for (const r of rows) {
      // 「行の途中で改ページしない」行は、収まらなければ次の段へ (Word の既定では行の途中で分ける)
      if (this.y + r.h > this.colBottom[this.col] + 0.5 && !this.atTop && r.cantSplit && r.h <= this.colHeight) next();
      if (this.y + r.h <= this.colBottom[this.col] + 0.5) {
        const items: PageItem[] = [];
        this.emitRow(r, this.y, items, merged);
        shift(items, dx(), 0);
        this.addUnit(items, r.h, 0);
        this.y += r.h;
        continue;
      }
      // 段に収まらない行は、段の下端で分ける (下端にかかる行は次の段へ)
      const contents = r.cells.flatMap((c) => c.items);
      let from = 0;
      for (let guard = 0; from < r.h - 0.5 && guard < 200; guard++) {
        const fit = this.colBottom[this.col] - this.y;
        let end = Math.min(r.h, from + Math.max(0, fit));
        const crossing = contents.filter((it) => it.y >= from && it.y < end && it.y + it.h > end);
        if (crossing.length) end = Math.min(...crossing.map((it) => it.y));
        if (end <= from + 0.5) {
          if (!this.atTop) {
            next();
            continue;
          }
          // 段の先頭でも1行も入らない: 段の高さで切る
          end = Math.min(r.h, from + Math.max(this.colHeight, 1));
        }
        const items: PageItem[] = [];
        for (const c of r.cells) items.push({ type: "cell", x: c.x, y: this.y, w: c.w, h: end - from });
        for (const it of contents) if (it.y >= from && it.y < end) items.push({ ...it, y: this.y + (it.y - from) });
        shift(items, dx(), 0);
        this.addUnit(items, end - from, 0);
        this.y += end - from;
        from = end;
        if (from < r.h - 0.5) next();
      }
      merged = new Map();
    }
    this.prev = undefined;
  }

  // ---- 全体 ------------------------------------------------------------------

  run(): LayoutResult {
    const { model } = this;
    let section = 0;
    this.startSection(0, true);
    const blocks = model.body;
    let keepCheckedUntil = -1;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      const s = b.kind === "para" ? model.paras[b.index].section : model.tables[b.table].section;
      while (section < s && section < model.sections.length - 1) this.startSection(++section, false);
      if (b.kind === "table") {
        this.placeTable(model.tables[b.table]);
        continue;
      }
      const p = model.paras[b.index];
      if (i > keepCheckedUntil && this.exists(p) && paraPropsAt(p, this.t).keepNext) keepCheckedUntil = this.keepWithNext(blocks, i);
      this.placePara(p);
    }
    // 中身の無いセクション (文書の最後の連続のセクション区切りなど): 段の高さをそろえる
    while (section < model.sections.length - 1) this.startSection(++section, false);
    // 文末脚注は文書の最後に置く
    if (this.endnotes.length) {
      const col = this.cols[this.col];
      for (const r of this.endnotes) {
        for (const l of this.noteLines(r, col.w)) {
          if (this.y + l.h > this.colBottom[this.col] && !this.atTop) this.advanceColumn();
          const c = this.cols[this.col];
          this.addUnit([{ type: "note", x: c.x, y: this.y, w: c.w, h: l.h, fill: l.fill }], l.h, 0);
          this.y += l.h;
        }
      }
    }
    this.finishPage();
    // 末尾の空ページ (最後の段落の直後の改ページ等) は除く
    const pages = [...this.pages];
    while (pages.length > 1 && pages[pages.length - 1].items.length === 0) pages.pop();
    return { pages, visibleChars: this.visibleChars, hintBreaks: this.hintBreaks };
  }
}

/** 時刻 t (ミリ秒) の文書を模式的なページに割り付ける */
export function layoutDocument(model: DocxLayoutModel, t: number, opts: LayoutOptions = {}): LayoutResult {
  return new Layouter(model, t, opts).run();
}
