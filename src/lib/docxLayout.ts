/**
 * docxLayout.ts
 *
 * 本文 (word/document.xml) を「段落」の列に変換し、段落内の各テキスト断片が
 * いつ挿入され (w:ins)、いつ削除されたか (w:del) を保持するモデルを作る。
 * このモデルから、任意の時刻 T における文書の状態 (T 以前の変更は反映し、
 * T より後の変更は元に戻した状態) を復元できる。
 *
 * あわせて、模式的なページ割り付け (pageLayout.ts) に使う情報を読み取る:
 *  - セクションごとの用紙・余白・文字数と行数 (docGrid)・段組み・縦書き・ヘッダーとフッターの高さ
 *  - 段落の書式 (スタイルの継承を解決したもの): 間隔・行間・インデント・配置・ページ送りの規則・文字の大きさ
 *  - 表の構造 (列の幅・セル・結合・行の高さ・行の挿入と削除)
 *  - 図の大きさと浮動配置、脚注・文末脚注の長さ、Word が保存時に改ページした位置 (w:lastRenderedPageBreak)
 *  - 書式の変更の履歴 (w:pPrChange / w:sectPrChange)。過去の時点の版は元の書式で割り付ける
 *
 * 制約:
 *  - 本物のレイアウトエンジンではないため、ページ割りは近似である。
 *  - テキストボックス内の文章は図として扱い、中身は解析しない。
 *  - w:moveFrom / w:moveTo (移動) は表示状態の復元に使い、移動元と移動先は Word が付ける
 *    範囲名 (w:moveFromRangeStart / w:moveToRangeStart の w:name) で対応づける。
 */

import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { t } from "./i18n";
import type { DocxInput } from "./input";

type XmlNode = Record<string, unknown>;

export interface RevRef {
  /** w:id。同じ id の断片は同じ変更に属する */
  id: string;
  /** 日時が無い (個人情報削除等) 場合は undefined */
  date?: Date;
  author?: string;
  /** 移動 (w:moveFrom / w:moveTo) による変更か */
  move: boolean;
  /** 移動の範囲名。同じ名前の moveFrom と moveTo が1つの移動の移動元と移動先 */
  moveName?: string;
}

export interface TextSeg {
  kind: "text";
  /** 本文 (移動・並べ替えの判定用) */
  text: string;
  /** 文字数 (集計用) */
  chars: number;
  /** 表示幅 (全角=1, 半角=0.5) */
  width: number;
  /** 文字の大きさ (pt)。スタイルと文字の書式から求めたもの */
  sizePt?: number;
  /** 隠し文字 (w:vanish)。ページには描かない */
  hidden?: boolean;
  ins?: RevRef;
  del?: RevRef;
}

export interface BreakSeg {
  kind: "lineBreak" | "pageBreak";
  /** 段区切り (w:br w:type="column")。段組みでなければ改ページと同じ */
  column?: boolean;
  ins?: RevRef;
  del?: RevRef;
}

/** 浮動配置の図 (wp:anchor) の位置と文字の回り込み */
export interface FloatSpec {
  /** 横位置の基準 (page / margin / column / character / leftMargin / rightMargin / insideMargin / outsideMargin) */
  relH: string;
  /** 基準からのオフセット (twip)。配置 (alignH) の指定が無いとき */
  offH?: number;
  /** 配置 (left / center / right / inside / outside) */
  alignH?: string;
  /** 縦位置の基準 (page / margin / paragraph / line / topMargin / bottomMargin) */
  relV: string;
  offV?: number;
  /** 配置 (top / center / bottom / inside / outside) */
  alignV?: string;
  /** 文字の回り込み: なし・前面・背面 (none)、上下 (topBottom)、四角・外周・内部 (square) */
  wrap: "none" | "topBottom" | "square";
  /** 文字との間隔 (twip) */
  dist: { t: number; b: number; l: number; r: number };
}

export interface FigureSeg {
  kind: "figure";
  /** 図の高さ (twip)。不明な場合は既定値 */
  heightTwips: number;
  /** 図の幅 (twip)。不明な場合は undefined (本文の幅の 70% で描く) */
  widthTwips?: number;
  /** 浮動配置の図。undefined なら行内の図 */
  float?: FloatSpec;
  ins?: RevRef;
  del?: RevRef;
}

export type Seg = TextSeg | BreakSeg | FigureSeg;

/** 段落の書式 (スタイルの継承を解決したもの)。長さは twip */
export interface ParaProps {
  styleId?: string;
  /** 見出しの階層 (0 = 見出し1)。見出しでなければ undefined */
  headingLevel?: number;
  spaceBefore: number;
  spaceAfter: number;
  /** 行単位の段落前後の間隔 (1/100 行)。指定があれば spaceBefore / spaceAfter より優先する */
  spaceBeforeLines?: number;
  spaceAfterLines?: number;
  /** 同じスタイルの段落の間の間隔を詰める */
  contextualSpacing: boolean;
  /** 行間。auto は 240 = 1行の倍数、exact / atLeast は twip */
  line: number;
  lineRule: "auto" | "exact" | "atLeast";
  /** 文書のグリッド (行送り・文字送り) に合わせる */
  snapToGrid: boolean;
  indLeft: number;
  indRight: number;
  /** 1行目のインデント。負の値はぶら下げ */
  indFirst: number;
  /** 字数で指定したインデント (1/100 字)。指定があれば twip の値より優先する */
  indLeftChars?: number;
  indRightChars?: number;
  indFirstChars?: number;
  jc: "left" | "center" | "right" | "both";
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  pageBreakBefore: boolean;
  /** 段落の文字の大きさ (pt)。文字ごとの指定は TextSeg.sizePt */
  sizePt: number;
  /** 1行の高さ ÷ 文字の大きさ (フォントから見積もった値) */
  lineFactor: number;
}

export interface NoteRef {
  /** 参照の直後の断片の位置 (Para.segs の添字) */
  segIndex: number;
  note: "footnote" | "endnote";
  id: string;
  ins?: RevRef;
  del?: RevRef;
}

export interface Para {
  index: number;
  kind: "body" | "heading" | "table";
  /** 見出しの階層 (0 = 見出し1)。見出しでなければ undefined */
  headingLevel?: number;
  /** 表の中の段落なら、その表 (最も外側の表) の通し番号 */
  tableId?: number;
  segs: Seg[];
  /** 段落記号そのものの挿入/削除 */
  markIns?: RevRef;
  markDel?: RevRef;
  pageBreakBefore: boolean;
  /** この段落の後でセクションが改ページされる */
  sectionBreakAfter: boolean;
  /**
   * この段落で始まるブックマーク (w:bookmarkStart の w:name)。段落と段落の間にあるものは次の段落に付ける。
   * Word が自動で付ける _GoBack は除く
   */
  bookmarks?: string[];
  /** Word が段落に付ける識別子 (w14:paraId、8桁の16進数)。Word for the web の見出しリンクに使う */
  paraId?: string;
  /** 段落が属するセクション (DocxLayoutModel.sections の添字) */
  section: number;
  /** 段落の書式 (現在) */
  props: ParaProps;
  /** 書式の変更 (w:pPrChange) の前の書式と、変更の日時。日時の無い変更は記録しない */
  propsBefore?: ParaProps;
  propsChangedAt?: Date;
  /** Word が最後に保存したとき、この位置 (segs の添字) の断片から新しいページ (段) が始まった */
  renderedBreaks?: number[];
  /** 脚注・文末脚注の参照 */
  noteRefs?: NoteRef[];
}

/** セクションの設定。長さは twip */
export interface SectionProps {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  gutter: number;
  headerDistance: number;
  footerDistance: number;
  /** 文書のグリッド: default (なし) / lines (行数のみ) / linesAndChars (文字数と行数) / snapToChars (文字数) */
  gridType: "default" | "lines" | "linesAndChars" | "snapToChars";
  /** 行送り */
  linePitch: number;
  /** 文字送りの調整 (1/4096 pt) */
  charSpace: number;
  /** 段組み。widths は段ごとの幅と、その段の後ろの段間 (均等でない場合) */
  columns: { num: number; space: number; separator: boolean; widths?: { width: number; space: number }[] };
  /** セクションの始まり方 */
  start: "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
  /** 縦書き (w:textDirection が tbRl など) */
  vertical: boolean;
  /** ヘッダー・フッターの本文の高さの見積もり */
  headerHeight: number;
  footerHeight: number;
}

export interface DocSection {
  props: SectionProps;
  /** 設定の変更 (w:sectPrChange) の前の設定と、変更の日時 */
  before?: SectionProps;
  changedAt?: Date;
}

export type BlockRef = { kind: "para"; index: number } | { kind: "table"; table: number };

export interface TableCell {
  /** 横に結合した列の数 */
  gridSpan: number;
  /** 縦の結合: restart (結合の先頭) / continue (上のセルと結合) */
  vMerge?: "restart" | "continue";
  /** セルの幅 (twip)。列の幅 (grid) が無いときに使う */
  width?: number;
  /** セルの余白 (twip) */
  margin?: { top?: number; bottom?: number; left?: number; right?: number };
  blocks: BlockRef[];
}

export interface TableRow {
  /** 行の挿入/削除 (w:trPr の w:ins / w:del) */
  ins?: RevRef;
  del?: RevRef;
  /** 行の高さの指定 (twip) と規則 */
  height?: number;
  heightRule: "auto" | "atLeast" | "exact";
  /** 行の途中で改ページしない */
  cantSplit: boolean;
  /** 見出し行 (ページごとに繰り返す) */
  header: boolean;
  /** 行の先頭で飛ばす列の数 (w:gridBefore) */
  gridBefore: number;
  cells: TableCell[];
}

export interface TableSpec {
  /** 段落の Para.tableId と同じ、最も外側の表の通し番号 */
  unitId: number;
  /** 表が属するセクション */
  section: number;
  /** 列の幅 (twip) */
  grid: number[];
  /** 表の幅の指定 (twip。pct の場合は undefined) */
  width?: number;
  indent: number;
  jc: "left" | "center" | "right";
  /** 既定のセルの余白 (twip) */
  cellMargin: { top: number; bottom: number; left: number; right: number };
  rows: TableRow[];
}

/** 脚注・文末脚注の長さ */
export interface NoteInfo {
  /** 本文の表示幅 (全角=1) */
  width: number;
  /** 段落の数 */
  paras: number;
}

export interface PageGeometry {
  pageWidthTwips: number;
  pageHeightTwips: number;
  marginTopTwips: number;
  marginBottomTwips: number;
  marginLeftTwips: number;
  marginRightTwips: number;
  /** 本文の行送り (twip) */
  lineHeightTwips: number;
  /** 本文の文字サイズ (pt) */
  fontSizePt: number;
}

export interface DocxLayoutModel {
  paras: Para[];
  /** 最後のセクションの用紙・余白・行送り (後方互換のため。割り付けには sections を使う) */
  geometry: PageGeometry;
  /** Word が最後に保存したときのページ数 (docProps/app.xml の Pages)。不明なら undefined */
  savedPageCount?: number;
  /** セクション (文書順) */
  sections: DocSection[];
  /** 本文の段落と表 (文書順)。表の中の段落は TableSpec の中にある */
  body: BlockRef[];
  /** 表 (入れ子の表を含む) */
  tables: TableSpec[];
  /** 脚注・文末脚注 (id ごと) */
  footnotes: Map<string, NoteInfo>;
  endnotes: Map<string, NoteInfo>;
  /** 脚注の文字の大きさ (pt) */
  noteSizePt: number;
  /** 既定のタブの間隔 (twip) */
  defaultTabStop: number;
  /** 文字送りの基準になる文字の大きさ (pt、既定の段落スタイルの大きさ) */
  baseSizePt: number;
}

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

function tagOf(node: XmlNode): string | undefined {
  return Object.keys(node).find((k) => k !== ":@" && k !== "#text");
}

function childrenOf(node: XmlNode, tag: string): XmlNode[] {
  const c = node[tag];
  return Array.isArray(c) ? (c as XmlNode[]) : [];
}

function attr(node: XmlNode, name: string): string | undefined {
  const attrs = node[":@"] as Record<string, unknown> | undefined;
  const v = attrs?.[`@_${name}`];
  return v === undefined ? undefined : String(v);
}

function findChild(nodes: XmlNode[], tag: string): XmlNode | undefined {
  return nodes.find((n) => tagOf(n) === tag);
}

/** 深さ優先で最初に見つかった tag の要素を返す */
function findDeep(nodes: XmlNode[], tag: string): XmlNode | undefined {
  for (const n of nodes) {
    const t = tagOf(n);
    if (t === undefined) continue;
    if (t === tag) return n;
    const found = findDeep(childrenOf(n, t), tag);
    if (found) return found;
  }
  return undefined;
}

function textOf(nodes: XmlNode[]): string {
  let s = "";
  for (const n of nodes) {
    if ("#text" in n) s += String(n["#text"]);
  }
  return s;
}

/** 数値の属性 (無い・数値でなければ undefined) */
function numAttr(node: XmlNode | undefined, name: string): number | undefined {
  const v = node ? Number(attr(node, name)) : NaN;
  return Number.isFinite(v) ? v : undefined;
}

/** オン/オフの値 ("true" / "1" / "on" ならオン) */
function isOn(v: string | undefined): boolean {
  return v !== undefined && !["false", "0", "off", "none"].includes(v);
}

/** オン/オフの要素 (w:keepNext など。w:val が無ければオン) */
function onOff(node: XmlNode | undefined): boolean | undefined {
  if (!node) return undefined;
  return isOn(attr(node, "val") ?? "true");
}

/** 全角相当の文字は幅1、それ以外は幅0.5 として表示幅を見積もる */
export function displayWidth(text: string): number {
  let w = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    w += cp >= 0x1100 ? 1 : 0.5;
  }
  return w;
}

// ---------------------------------------------------------------------------
// 書式 (スタイルの継承)
// ---------------------------------------------------------------------------

/** 書式のうち、指定されたものだけ (継承の途中の値) */
type PPrRaw = Partial<Omit<ParaProps, "headingLevel" | "sizePt" | "lineFactor" | "styleId">> & {
  outlineLvl?: number;
  numId?: string;
  ilvl?: number;
};
interface FontsRaw {
  ascii?: string;
  eastAsia?: string;
  asciiTheme?: string;
  eastAsiaTheme?: string;
}
interface RPrRaw {
  sizePt?: number;
  vanish?: boolean;
  fonts?: FontsRaw;
}

interface StyleInfo {
  name?: string;
  type?: string;
  isDefault: boolean;
  basedOn?: string;
  outlineLvl?: number;
  ppr: PPrRaw;
  rpr: RPrRaw;
}

function parsePPr(pPr: XmlNode | undefined): PPrRaw {
  const out: PPrRaw = {};
  if (!pPr) return out;
  const kids = childrenOf(pPr, "pPr");
  const spacing = findChild(kids, "spacing");
  if (spacing) {
    const set = <K extends keyof PPrRaw>(k: K, v: PPrRaw[K] | undefined) => {
      if (v !== undefined) out[k] = v;
    };
    set("spaceBefore", numAttr(spacing, "before"));
    set("spaceAfter", numAttr(spacing, "after"));
    set("spaceBeforeLines", numAttr(spacing, "beforeLines"));
    set("spaceAfterLines", numAttr(spacing, "afterLines"));
    // 自動 (HTML の段落間隔) は 14pt とみなす
    if (isOn(attr(spacing, "beforeAutospacing"))) out.spaceBefore = 280;
    if (isOn(attr(spacing, "afterAutospacing"))) out.spaceAfter = 280;
    const line = numAttr(spacing, "line");
    if (line !== undefined && line > 0) {
      out.line = line;
      const rule = attr(spacing, "lineRule");
      out.lineRule = rule === "exact" ? "exact" : rule === "atLeast" ? "atLeast" : "auto";
    }
  }
  const cs = findChild(kids, "contextualSpacing");
  if (cs) out.contextualSpacing = onOff(cs);
  const snap = findChild(kids, "snapToGrid");
  if (snap) out.snapToGrid = onOff(snap);
  const ind = findChild(kids, "ind");
  if (ind) {
    const left = numAttr(ind, "left") ?? numAttr(ind, "start");
    const right = numAttr(ind, "right") ?? numAttr(ind, "end");
    if (left !== undefined) out.indLeft = left;
    if (right !== undefined) out.indRight = right;
    const leftChars = numAttr(ind, "leftChars") ?? numAttr(ind, "startChars");
    const rightChars = numAttr(ind, "rightChars") ?? numAttr(ind, "endChars");
    if (leftChars !== undefined) out.indLeftChars = leftChars;
    if (rightChars !== undefined) out.indRightChars = rightChars;
    const hanging = numAttr(ind, "hanging");
    const firstLine = numAttr(ind, "firstLine");
    if (hanging !== undefined) out.indFirst = -hanging;
    else if (firstLine !== undefined) out.indFirst = firstLine;
    const hangingChars = numAttr(ind, "hangingChars");
    const firstLineChars = numAttr(ind, "firstLineChars");
    if (hangingChars !== undefined) out.indFirstChars = -hangingChars;
    else if (firstLineChars !== undefined) out.indFirstChars = firstLineChars;
  }
  const jc = findChild(kids, "jc");
  if (jc) {
    const v = attr(jc, "val");
    out.jc = v === "center" ? "center" : v === "right" || v === "end" ? "right" : v === "both" || v === "distribute" ? "both" : "left";
  }
  for (const [tag, key] of [
    ["keepNext", "keepNext"],
    ["keepLines", "keepLines"],
    ["widowControl", "widowControl"],
    ["pageBreakBefore", "pageBreakBefore"],
  ] as const) {
    const n = findChild(kids, tag);
    if (n) out[key] = onOff(n);
  }
  const outline = findChild(kids, "outlineLvl");
  if (outline) out.outlineLvl = numAttr(outline, "val");
  const numPr = findChild(kids, "numPr");
  if (numPr) {
    const nk = childrenOf(numPr, "numPr");
    const numId = findChild(nk, "numId");
    const ilvl = findChild(nk, "ilvl");
    if (numId) out.numId = attr(numId, "val");
    if (ilvl) out.ilvl = numAttr(ilvl, "val");
  }
  return out;
}

function parseRPr(rPr: XmlNode | undefined): RPrRaw {
  const out: RPrRaw = {};
  if (!rPr) return out;
  const kids = childrenOf(rPr, "rPr");
  const sz = findChild(kids, "sz");
  const v = numAttr(sz, "val");
  if (v !== undefined && v > 0) out.sizePt = v / 2;
  const vanish = findChild(kids, "vanish");
  if (vanish) out.vanish = onOff(vanish);
  const fonts = findChild(kids, "rFonts");
  if (fonts) {
    out.fonts = {
      ascii: attr(fonts, "ascii"),
      eastAsia: attr(fonts, "eastAsia"),
      asciiTheme: attr(fonts, "asciiTheme"),
      eastAsiaTheme: attr(fonts, "eastAsiaTheme"),
    };
  }
  return out;
}

function mergeFonts(a: FontsRaw | undefined, b: FontsRaw | undefined): FontsRaw | undefined {
  if (!b) return a;
  const out: FontsRaw = { ...a };
  // 文字の書式で直接フォントを指定すると、テーマのフォントの指定は外れる
  if (b.ascii) Object.assign(out, { ascii: b.ascii, asciiTheme: undefined });
  if (b.asciiTheme) out.asciiTheme = b.asciiTheme;
  if (b.eastAsia) Object.assign(out, { eastAsia: b.eastAsia, eastAsiaTheme: undefined });
  if (b.eastAsiaTheme) out.eastAsiaTheme = b.eastAsiaTheme;
  return out;
}

function mergeRPr(a: RPrRaw, b: RPrRaw): RPrRaw {
  return {
    sizePt: b.sizePt ?? a.sizePt,
    vanish: b.vanish ?? a.vanish,
    fonts: mergeFonts(a.fonts, b.fonts),
  };
}

function parseStyles(root: XmlNode[] | undefined): Map<string, StyleInfo> {
  const map = new Map<string, StyleInfo>();
  if (!root) return map;
  const stylesNode = root.find((n) => tagOf(n) === "styles");
  for (const s of stylesNode ? childrenOf(stylesNode, "styles") : []) {
    if (tagOf(s) !== "style") continue;
    const id = attr(s, "styleId");
    if (!id) continue;
    const kids = childrenOf(s, "style");
    const name = findChild(kids, "name");
    const basedOn = findChild(kids, "basedOn");
    const ppr = parsePPr(findChild(kids, "pPr"));
    map.set(id, {
      name: name ? attr(name, "val") : undefined,
      type: attr(s, "type"),
      isDefault: ["1", "true", "on"].includes(String(attr(s, "default") ?? "")),
      basedOn: basedOn ? attr(basedOn, "val") : undefined,
      outlineLvl: ppr.outlineLvl,
      ppr,
      rpr: parseRPr(findChild(kids, "rPr")),
    });
  }
  return map;
}

/** スタイルから見出しの階層を求める (basedOn をたどる) */
function headingLevelOf(styleId: string | undefined, styles: Map<string, StyleInfo>): number | undefined {
  for (let id = styleId, depth = 0; id && depth < 10; depth++) {
    const s = styles.get(id);
    if (!s) break;
    if (s.outlineLvl !== undefined && s.outlineLvl >= 0 && s.outlineLvl <= 8) return s.outlineLvl;
    const m = s.name?.match(/^heading\s*(\d)$/i);
    if (m) return Number(m[1]) - 1;
    id = s.basedOn;
  }
  return undefined;
}

/** スタイルの継承の列 (根元のスタイルから順に) */
function styleChain(styleId: string | undefined, styles: Map<string, StyleInfo>): StyleInfo[] {
  const chain: StyleInfo[] = [];
  const seen = new Set<string>();
  for (let id = styleId; id && !seen.has(id) && chain.length < 12; ) {
    seen.add(id);
    const s = styles.get(id);
    if (!s) break;
    chain.unshift(s);
    id = s.basedOn;
  }
  return chain;
}

/**
 * フォントごとの、1行の高さ ÷ 文字の大きさ (Word の「1行」の行間の見積もり)。
 * フォントの寸法 (アセント・ディセント) は環境によって違うため、よく使われるフォントの概算値を持つ
 */
const LINE_FACTORS: [RegExp, number][] = [
  [/^(ｍｓ|ms)p?(明朝|mincho|ゴシック|gothic)/i, 1.0],
  [/^(游|yu)(明朝|mincho|ゴシック|gothic)/i, 1.7],
  [/^(メイリオ|meiryo)/i, 1.5],
  [/^(ヒラギノ|hiragino)/i, 1.5],
  [/^(noto|源ノ|source han)/i, 1.45],
  [/^biz ?ud/i, 1.4],
  [/^calibri/i, 1.22],
  [/^(century|cambria)/i, 1.17],
  [/^(times new roman|arial|helvetica)/i, 1.15],
];
const DEFAULT_LINE_FACTOR = 1.2;

function lineFactorOf(font: string | undefined): number | undefined {
  if (!font) return undefined;
  const f = font.trim();
  return LINE_FACTORS.find(([re]) => re.test(f))?.[1];
}

interface ThemeFonts {
  majorLatin?: string;
  majorEastAsia?: string;
  minorLatin?: string;
  minorEastAsia?: string;
}

/** テーマのフォント (word/theme/theme1.xml)。東アジアのフォントは ea、空なら日本語 (Jpan) の指定 */
function parseThemeFonts(xml: string | undefined): ThemeFonts {
  if (!xml) return {};
  const block = (name: string) => xml.match(new RegExp(`<a:${name}>([\\s\\S]*?)</a:${name}>`))?.[1] ?? "";
  const pick = (b: string) => {
    const latin = b.match(/<a:latin\b[^>]*typeface="([^"]*)"/)?.[1];
    const ea = b.match(/<a:ea\b[^>]*typeface="([^"]*)"/)?.[1];
    const jpan = b.match(/<a:font\b[^>]*script="Jpan"[^>]*typeface="([^"]*)"/)?.[1];
    return { latin: latin || undefined, ea: ea || jpan || undefined };
  };
  const major = pick(block("majorFont"));
  const minor = pick(block("minorFont"));
  return { majorLatin: major.latin, majorEastAsia: major.ea, minorLatin: minor.latin, minorEastAsia: minor.ea };
}

function fontsLineFactor(fonts: FontsRaw | undefined, theme: ThemeFonts): number {
  const resolve = (name: string | undefined, themeName: string | undefined) => {
    if (themeName) {
      if (/^major/i.test(themeName)) return /EastAsia/i.test(themeName) ? theme.majorEastAsia : theme.majorLatin;
      return /EastAsia/i.test(themeName) ? theme.minorEastAsia : theme.minorLatin;
    }
    return name;
  };
  const ascii = lineFactorOf(resolve(fonts?.ascii, fonts?.asciiTheme ?? (fonts?.ascii ? undefined : "minorHAnsi")));
  const ea = lineFactorOf(resolve(fonts?.eastAsia, fonts?.eastAsiaTheme ?? (fonts?.eastAsia ? undefined : "minorEastAsia")));
  const known = [ascii, ea].filter((v): v is number => v !== undefined);
  return known.length ? Math.max(...known) : DEFAULT_LINE_FACTOR;
}

/** 段落番号の字下げ (word/numbering.xml): numId → ilvl → 書式 */
type Numbering = Map<string, Map<number, PPrRaw>>;

function parseNumbering(root: XmlNode[] | undefined): Numbering {
  const out: Numbering = new Map();
  const numberingNode = root?.find((n) => tagOf(n) === "numbering");
  if (!numberingNode) return out;
  const abstract = new Map<string, Map<number, PPrRaw>>();
  const kids = childrenOf(numberingNode, "numbering");
  for (const a of kids) {
    if (tagOf(a) !== "abstractNum") continue;
    const levels = new Map<number, PPrRaw>();
    for (const lvl of childrenOf(a, "abstractNum")) {
      if (tagOf(lvl) !== "lvl") continue;
      levels.set(numAttr(lvl, "ilvl") ?? 0, parsePPr(findChild(childrenOf(lvl, "lvl"), "pPr")));
    }
    abstract.set(attr(a, "abstractNumId") ?? "", levels);
  }
  for (const n of kids) {
    if (tagOf(n) !== "num") continue;
    const ref = findChild(childrenOf(n, "num"), "abstractNumId");
    const levels = ref ? abstract.get(attr(ref, "val") ?? "") : undefined;
    if (levels) out.set(attr(n, "numId") ?? "", levels);
  }
  return out;
}

interface StyleContext {
  styles: Map<string, StyleInfo>;
  numbering: Numbering;
  theme: ThemeFonts;
  defaultPPr: PPrRaw;
  defaultRPr: RPrRaw;
  /** 既定の段落スタイル */
  defaultParaStyle?: string;
}

const BASE_PROPS: Omit<ParaProps, "sizePt" | "lineFactor"> = {
  spaceBefore: 0,
  spaceAfter: 0,
  contextualSpacing: false,
  line: 240,
  lineRule: "auto",
  snapToGrid: true,
  indLeft: 0,
  indRight: 0,
  indFirst: 0,
  jc: "left",
  keepNext: false,
  keepLines: false,
  widowControl: true,
  pageBreakBefore: false,
};

/** 段落の書式を解決する: 既定値 → docDefaults → スタイル (basedOn の根元から) → 段落番号 → 直接の書式 */
function resolveParaProps(direct: PPrRaw, styleId: string | undefined, ctx: StyleContext): { props: ParaProps; rpr: RPrRaw } {
  const sid = styleId ?? ctx.defaultParaStyle;
  const chain = styleChain(sid, ctx.styles);
  let ppr: PPrRaw = { ...ctx.defaultPPr };
  let rpr: RPrRaw = { ...ctx.defaultRPr };
  for (const s of chain) {
    ppr = { ...ppr, ...s.ppr };
    rpr = mergeRPr(rpr, s.rpr);
  }
  const numId = direct.numId ?? ppr.numId;
  const ilvl = direct.ilvl ?? ppr.ilvl ?? 0;
  const lvl = numId && numId !== "0" ? ctx.numbering.get(numId)?.get(ilvl) : undefined;
  if (lvl) ppr = { ...ppr, ...lvl };
  ppr = { ...ppr, ...direct };
  let level = headingLevelOf(sid, ctx.styles);
  if (direct.outlineLvl !== undefined) level = direct.outlineLvl >= 0 && direct.outlineLvl <= 8 ? direct.outlineLvl : undefined;
  const { outlineLvl: _o, numId: _n, ilvl: _i, ...rest } = ppr;
  const props: ParaProps = {
    ...BASE_PROPS,
    ...(Object.fromEntries(Object.entries(rest).filter(([, v]) => v !== undefined)) as Partial<ParaProps>),
    styleId: sid,
    headingLevel: level,
    sizePt: rpr.sizePt ?? 10.5,
    lineFactor: fontsLineFactor(rpr.fonts, ctx.theme),
  };
  return { props, rpr };
}

// ---------------------------------------------------------------------------
// 本文の走査
// ---------------------------------------------------------------------------

/** 文書順に走査しながら、現在どの移動範囲の中にいるかを覚えておく */
interface ParseState {
  moveFromName?: string;
  moveToName?: string;
  /** いま読んでいる段落で始まったブックマーク / 段落の外で始まり、次の段落に付けるブックマーク */
  paraBookmarks?: string[];
  pendingBookmarks?: string[];
  /** いま読んでいる段落の、Word が保存時に改ページした位置と脚注の参照 */
  renderedBreaks?: number[];
  noteRefs?: NoteRef[];
}

/** 段落に付けるブックマークの名前 (Word が自動で付ける _GoBack は除く) */
function bookmarkName(node: XmlNode): string | undefined {
  const name = attr(node, "name");
  return name && name !== "_GoBack" ? name : undefined;
}

function revRef(node: XmlNode, move: boolean, moveName?: string): RevRef {
  const dateStr = attr(node, "date");
  return {
    id: attr(node, "id") ?? "",
    date: dateStr ? new Date(dateStr) : undefined,
    author: attr(node, "author"),
    move,
    moveName: move ? moveName : undefined,
  };
}

/** 移動範囲の開始・終了の印を処理する。処理した場合は true */
function trackMoveRange(tag: string, node: XmlNode, state: ParseState): boolean {
  switch (tag) {
    case "moveFromRangeStart":
      state.moveFromName = attr(node, "name") ?? attr(node, "id");
      return true;
    case "moveFromRangeEnd":
      state.moveFromName = undefined;
      return true;
    case "moveToRangeStart":
      state.moveToName = attr(node, "name") ?? attr(node, "id");
      return true;
    case "moveToRangeEnd":
      state.moveToName = undefined;
      return true;
  }
  return false;
}

const EMU_PER_TWIP = 635;
const DEFAULT_FIGURE_HEIGHT_TWIPS = 2880; // 2インチ

/** "12pt" / "1.5in" / "2cm" などの長さ (VML の style) を twip にする */
function cssLength(v: string | undefined): number | undefined {
  const m = v?.trim().match(/^(-?\d+(?:\.\d+)?)(pt|in|cm|mm|px)?$/);
  if (!m) return undefined;
  const n = Number(m[1]);
  const per = { pt: 20, in: 1440, cm: 1440 / 2.54, mm: 144 / 2.54, px: 15 }[(m[2] ?? "px") as "pt"];
  return n * per;
}

/** 浮動配置の図の位置と回り込み (wp:anchor) */
function anchorFloat(anchor: XmlNode): FloatSpec {
  const kids = childrenOf(anchor, "anchor");
  const pos = (tag: "positionH" | "positionV") => {
    const n = findChild(kids, tag);
    const nk = n ? childrenOf(n, tag) : [];
    const off = findChild(nk, "posOffset");
    const align = findChild(nk, "align");
    return {
      rel: (n && attr(n, "relativeFrom")) ?? (tag === "positionH" ? "column" : "paragraph"),
      off: off ? Number(textOf(childrenOf(off, "posOffset"))) / EMU_PER_TWIP : undefined,
      align: align ? textOf(childrenOf(align, "align")).trim() : undefined,
    };
  };
  const h = pos("positionH");
  const v = pos("positionV");
  const wrapTag = kids.map(tagOf).find((t) => t?.startsWith("wrap"));
  const dist = (name: string) => (numAttr(anchor, name) ?? 0) / EMU_PER_TWIP;
  return {
    relH: h.rel,
    offH: h.off,
    alignH: h.align,
    relV: v.rel,
    offV: v.off,
    alignV: v.align,
    wrap: wrapTag === "wrapTopAndBottom" ? "topBottom" : wrapTag === "wrapNone" || !wrapTag ? "none" : "square",
    dist: { t: dist("distT"), b: dist("distB"), l: dist("distL"), r: dist("distR") },
  };
}

/** 図 (drawing / pict / object / AlternateContent) の大きさと配置 */
function figureSeg(nodes: XmlNode[]): Omit<FigureSeg, "kind" | "ins" | "del"> {
  const anchor = findDeep(nodes, "anchor");
  const extent = findDeep(nodes, "extent");
  const cx = numAttr(extent, "cx");
  const cy = numAttr(extent, "cy");
  if (cy !== undefined && cy > 0) {
    return {
      heightTwips: cy / EMU_PER_TWIP,
      widthTwips: cx !== undefined && cx > 0 ? cx / EMU_PER_TWIP : undefined,
      float: anchor ? anchorFloat(anchor) : undefined,
    };
  }
  // VML (v:shape など) は style 属性に大きさと配置を持つ
  const shape = ["shape", "rect", "roundrect", "oval", "group"].map((tag) => findDeep(nodes, tag)).find(Boolean);
  const style = shape ? attr(shape, "style") : undefined;
  if (style) {
    const css = new Map(
      style.split(";").map((d) => {
        const i = d.indexOf(":");
        return [d.slice(0, i).trim().toLowerCase(), d.slice(i + 1).trim()] as [string, string];
      })
    );
    const h = cssLength(css.get("height"));
    const w = cssLength(css.get("width"));
    if (h && h > 0) {
      const absolute = css.get("position") === "absolute";
      const wrap = findDeep(childrenOf(shape!, tagOf(shape!)!), "wrap");
      const wrapType = wrap ? attr(wrap, "type") : undefined;
      return {
        heightTwips: h,
        widthTwips: w,
        float: absolute
          ? {
              relH: css.get("mso-position-horizontal-relative") ?? "column",
              offH: cssLength(css.get("margin-left")),
              relV: css.get("mso-position-vertical-relative") ?? "paragraph",
              offV: cssLength(css.get("margin-top")),
              wrap: wrapType === "topAndBottom" ? "topBottom" : wrapType === "none" || !wrapType ? "none" : "square",
              dist: { t: 0, b: 0, l: 0, r: 0 },
            }
          : undefined,
      };
    }
  }
  return { heightTwips: DEFAULT_FIGURE_HEIGHT_TWIPS };
}

interface RevContext {
  ins?: RevRef;
  del?: RevRef;
  /** 文字の書式 (段落の書式に、文字スタイルと直接の書式を重ねたもの) */
  rpr?: RPrRaw;
}

/** 段落の子要素を走査し、テキスト断片等を segs に追加する */
function walkInline(nodes: XmlNode[], ctx: RevContext, segs: Seg[], state: ParseState, sc: StyleContext): void {
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined) continue;
    if (trackMoveRange(tag, node, state)) continue;
    const kids = childrenOf(node, tag);
    switch (tag) {
      case "ins":
        walkInline(kids, { ...ctx, ins: revRef(node, false) }, segs, state, sc);
        break;
      case "moveTo":
        walkInline(kids, { ...ctx, ins: revRef(node, true, state.moveToName) }, segs, state, sc);
        break;
      case "del":
        walkInline(kids, { ...ctx, del: revRef(node, false) }, segs, state, sc);
        break;
      case "moveFrom":
        walkInline(kids, { ...ctx, del: revRef(node, true, state.moveFromName) }, segs, state, sc);
        break;
      case "r": {
        // 文字の書式: 文字スタイル → 直接の書式
        const rPrNode = findChild(kids, "rPr");
        let rpr = ctx.rpr ?? {};
        if (rPrNode) {
          const rStyle = findChild(childrenOf(rPrNode, "rPr"), "rStyle");
          for (const s of styleChain(rStyle ? attr(rStyle, "val") : undefined, sc.styles)) rpr = mergeRPr(rpr, s.rpr);
          rpr = mergeRPr(rpr, parseRPr(rPrNode));
        }
        walkInline(kids, { ...ctx, rpr }, segs, state, sc);
        break;
      }
      case "t":
      case "delText": {
        const text = textOf(kids);
        if (text.length > 0) {
          const { rpr, ...rev } = ctx;
          const seg: TextSeg = { kind: "text", text, chars: [...text].length, width: displayWidth(text), ...rev };
          if (rpr?.sizePt !== undefined) seg.sizePt = rpr.sizePt;
          if (rpr?.vanish) seg.hidden = true;
          segs.push(seg);
        }
        break;
      }
      case "tab": {
        const { rpr: _r, ...rev } = ctx;
        segs.push({ kind: "text", text: "\t", chars: 1, width: 2, ...rev });
        break;
      }
      case "br":
      case "cr": {
        const { rpr: _r, ...rev } = ctx;
        const type = attr(node, "type");
        if (type === "column") segs.push({ kind: "lineBreak", column: true, ...rev });
        else segs.push({ kind: type === "page" ? "pageBreak" : "lineBreak", ...rev });
        break;
      }
      case "drawing":
      case "pict":
      case "object":
      case "AlternateContent": {
        const { rpr: _r, ...rev } = ctx;
        segs.push({ kind: "figure", ...figureSeg(kids), ...rev });
        break;
      }
      case "footnoteReference":
      case "endnoteReference": {
        const id = attr(node, "id");
        if (id !== undefined) {
          (state.noteRefs ??= []).push({
            segIndex: segs.length,
            note: tag === "footnoteReference" ? "footnote" : "endnote",
            id,
            ins: ctx.ins,
            del: ctx.del,
          });
        }
        break;
      }
      case "lastRenderedPageBreak":
        // 削除された文字の中の印は、最終状態の手がかりにならない
        if (!ctx.del) (state.renderedBreaks ??= []).push(segs.length);
        break;
      // 書式・フィールドコード等、本文として表示されないもの
      case "rPr":
      case "pPr":
      case "instrText":
      case "delInstrText":
      case "fldChar":
        break;
      case "bookmarkStart": {
        const name = bookmarkName(node);
        if (name) (state.paraBookmarks ??= []).push(name);
        break;
      }
      case "bookmarkEnd":
      case "proofErr":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
        break;
      default:
        // hyperlink, smartTag, sdt, sdtContent, fldSimple, customXml 等は中へ潜る
        walkInline(kids, ctx, segs, state, sc);
    }
  }
}

function sectPrBreaksPage(sectPr: XmlNode): boolean {
  const type = findChild(childrenOf(sectPr, "sectPr"), "type");
  const v = type ? attr(type, "val") : undefined;
  return v !== "continuous" && v !== "nextColumn";
}

function parseParagraph(
  node: XmlNode,
  index: number,
  tableId: number | undefined,
  section: number,
  sc: StyleContext,
  state: ParseState
): { para: Para; sectPr?: XmlNode } {
  const inTable = tableId !== undefined;
  const kids = childrenOf(node, "p");
  const pPr = findChild(kids, "pPr");
  const pKids = pPr ? childrenOf(pPr, "pPr") : [];
  const pStyle = findChild(pKids, "pStyle");
  const { props, rpr } = resolveParaProps(parsePPr(pPr), pStyle ? attr(pStyle, "val") : undefined, sc);
  const para: Para = {
    index,
    kind: inTable ? "table" : "body",
    tableId,
    segs: [],
    pageBreakBefore: props.pageBreakBefore,
    sectionBreakAfter: false,
    section,
    props,
  };
  const paraId = attr(node, "paraId");
  if (paraId && /^[0-9A-Fa-f]{8}$/.test(paraId)) para.paraId = paraId.toUpperCase();
  if (props.headingLevel !== undefined && !inTable) {
    para.kind = "heading";
    para.headingLevel = props.headingLevel;
  }
  // 書式の変更の前の書式 (日時の無い変更は「最初から反映済み」とみなすので持たない)
  const change = findChild(pKids, "pPrChange");
  const changeDate = change ? attr(change, "date") : undefined;
  if (change && changeDate) {
    const oldPPr = findChild(childrenOf(change, "pPrChange"), "pPr");
    const oldStyle = oldPPr ? findChild(childrenOf(oldPPr, "pPr"), "pStyle") : undefined;
    para.propsBefore = resolveParaProps(parsePPr(oldPPr), oldStyle ? attr(oldStyle, "val") : undefined, sc).props;
    para.propsChangedAt = new Date(changeDate);
  }
  const sectPr = findChild(pKids, "sectPr");
  if (sectPr && sectPrBreaksPage(sectPr)) para.sectionBreakAfter = true;
  // 段落記号の挿入/削除は pPr/rPr の中に記録される
  const markRPr = findChild(pKids, "rPr");
  if (markRPr) {
    const rKids = childrenOf(markRPr, "rPr");
    const ins = findChild(rKids, "ins") ?? findChild(rKids, "moveTo");
    const del = findChild(rKids, "del") ?? findChild(rKids, "moveFrom");
    if (ins) para.markIns = revRef(ins, tagOf(ins) === "moveTo", state.moveToName);
    if (del) para.markDel = revRef(del, tagOf(del) === "moveFrom", state.moveFromName);
  }
  state.paraBookmarks = state.pendingBookmarks ?? [];
  state.pendingBookmarks = undefined;
  Object.assign(state, { renderedBreaks: undefined, noteRefs: undefined });
  walkInline(kids, { rpr }, para.segs, state, sc);
  if (state.paraBookmarks.length) para.bookmarks = state.paraBookmarks;
  if (state.renderedBreaks?.length) para.renderedBreaks = state.renderedBreaks;
  if (state.noteRefs?.length) para.noteRefs = state.noteRefs;
  state.paraBookmarks = undefined;
  return { para, sectPr };
}

interface BlockContext {
  sc: StyleContext;
  out: Para[];
  tables: TableSpec[];
  tableCount: number;
  state: ParseState;
  /** いまのセクションの番号と、これまでに見つかったセクションの w:sectPr */
  section: number;
  sectPrs: XmlNode[];
}

/** 本文ブロック (body, tc, sdtContent 等) を走査して、段落と表を container に集める */
function walkBlocks(nodes: XmlNode[], tableId: number | undefined, ctx: BlockContext, container: BlockRef[]): XmlNode | undefined {
  let lastSectPr: XmlNode | undefined;
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined) continue;
    if (trackMoveRange(tag, node, ctx.state)) continue;
    if (tag === "bookmarkStart") {
      // 段落と段落の間で始まるブックマークは、次の段落に付ける
      const name = bookmarkName(node);
      if (name) (ctx.state.pendingBookmarks ??= []).push(name);
      continue;
    }
    if (tag === "p") {
      const { para, sectPr } = parseParagraph(node, ctx.out.length, tableId, ctx.section, ctx.sc, ctx.state);
      ctx.out.push(para);
      container.push({ kind: "para", index: para.index });
      // 段落の w:sectPr は、その段落で終わるセクションの設定 (表の中には現れない)
      if (sectPr && tableId === undefined) {
        ctx.sectPrs.push(sectPr);
        ctx.section++;
      }
    } else if (tag === "tbl") {
      // 入れ子の表は、段落の tableId では外側の表の一部として扱う
      const unitId = tableId ?? ctx.tableCount++;
      container.push({ kind: "table", table: parseTable(node, unitId, ctx) });
    } else if (tag === "sectPr") {
      lastSectPr = node;
    } else if (["sdt", "sdtContent", "customXml", "ins", "del", "moveTo", "moveFrom"].includes(tag)) {
      walkBlocks(childrenOf(node, tag), tableId, ctx, container);
    }
  }
  return lastSectPr;
}

/** 表の行 (w:tr) を、sdt などの入れ物の中も含めて文書順に集める */
function collectChildren(nodes: XmlNode[], want: string, out: XmlNode[] = []): XmlNode[] {
  for (const n of nodes) {
    const tag = tagOf(n);
    if (tag === want) out.push(n);
    else if (tag && ["sdt", "sdtContent", "customXml", "ins", "del", "moveTo", "moveFrom"].includes(tag)) {
      collectChildren(childrenOf(n, tag), want, out);
    }
  }
  return out;
}

function marginsOf(node: XmlNode | undefined): { top?: number; bottom?: number; left?: number; right?: number } | undefined {
  if (!node) return undefined;
  const kids = childrenOf(node, tagOf(node)!);
  const w = (tag: string, alt?: string) => {
    const n = findChild(kids, tag) ?? (alt ? findChild(kids, alt) : undefined);
    return n && attr(n, "type") !== "pct" ? numAttr(n, "w") : undefined;
  };
  return { top: w("top"), bottom: w("bottom"), left: w("left", "start"), right: w("right", "end") };
}

function parseTable(node: XmlNode, unitId: number, ctx: BlockContext): number {
  const kids = childrenOf(node, "tbl");
  const tblPr = findChild(kids, "tblPr");
  const prKids = tblPr ? childrenOf(tblPr, "tblPr") : [];
  const tblW = findChild(prKids, "tblW");
  const tblInd = findChild(prKids, "tblInd");
  const jc = findChild(prKids, "jc");
  const jcVal = jc ? attr(jc, "val") : undefined;
  const mar = marginsOf(findChild(prKids, "tblCellMar"));
  const grid = findChild(kids, "tblGrid");
  const spec: TableSpec = {
    unitId,
    section: ctx.section,
    grid: grid
      ? childrenOf(grid, "tblGrid")
          .filter((g) => tagOf(g) === "gridCol")
          .map((g) => numAttr(g, "w") ?? 0)
      : [],
    width: tblW && attr(tblW, "type") === "dxa" ? numAttr(tblW, "w") : undefined,
    indent: tblInd && attr(tblInd, "type") !== "pct" ? numAttr(tblInd, "w") ?? 0 : 0,
    jc: jcVal === "center" ? "center" : jcVal === "right" || jcVal === "end" ? "right" : "left",
    cellMargin: { top: mar?.top ?? 0, bottom: mar?.bottom ?? 0, left: mar?.left ?? 108, right: mar?.right ?? 108 },
    rows: [],
  };
  const id = ctx.tables.length;
  ctx.tables.push(spec);
  for (const tr of collectChildren(kids, "tr")) {
    const trKids = childrenOf(tr, "tr");
    const trPr = findChild(trKids, "trPr");
    const trPrKids = trPr ? childrenOf(trPr, "trPr") : [];
    const ins = findChild(trPrKids, "ins");
    const del = findChild(trPrKids, "del");
    const trHeight = findChild(trPrKids, "trHeight");
    const hRule = trHeight ? attr(trHeight, "hRule") : undefined;
    const gridBefore = findChild(trPrKids, "gridBefore");
    const row: TableRow = {
      ins: ins ? revRef(ins, false) : undefined,
      del: del ? revRef(del, false) : undefined,
      height: numAttr(trHeight, "val"),
      heightRule: !trHeight ? "auto" : hRule === "exact" ? "exact" : hRule === "auto" ? "auto" : "atLeast",
      cantSplit: !!onOff(findChild(trPrKids, "cantSplit")),
      header: !!onOff(findChild(trPrKids, "tblHeader")),
      gridBefore: numAttr(gridBefore, "val") ?? 0,
      cells: [],
    };
    for (const tc of collectChildren(trKids, "tc")) {
      const tcKids = childrenOf(tc, "tc");
      const tcPr = findChild(tcKids, "tcPr");
      const tcPrKids = tcPr ? childrenOf(tcPr, "tcPr") : [];
      const span = findChild(tcPrKids, "gridSpan");
      const vMerge = findChild(tcPrKids, "vMerge");
      const tcW = findChild(tcPrKids, "tcW");
      const cell: TableCell = {
        gridSpan: Math.max(1, numAttr(span, "val") ?? 1),
        vMerge: vMerge ? (attr(vMerge, "val") === "restart" ? "restart" : "continue") : undefined,
        width: tcW && attr(tcW, "type") === "dxa" ? numAttr(tcW, "w") : undefined,
        margin: marginsOf(findChild(tcPrKids, "tcMar")),
        blocks: [],
      };
      walkBlocks(tcKids, unitId, ctx, cell.blocks);
      row.cells.push(cell);
    }
    spec.rows.push(row);
  }
  return id;
}

// ---------------------------------------------------------------------------
// セクション
// ---------------------------------------------------------------------------

/** ヘッダー・フッター・脚注の本文の長さ */
function storyLength(root: XmlNode[]): { width: number; paras: number; figureHeight: number } {
  let width = 0;
  let paras = 0;
  let figureHeight = 0;
  const walk = (nodes: XmlNode[]) => {
    for (const n of nodes) {
      const tag = tagOf(n);
      if (tag === undefined) continue;
      if (tag === "p") paras++;
      if (tag === "t") width += displayWidth(textOf(childrenOf(n, tag)));
      if (tag === "drawing" || tag === "pict") {
        const f = figureSeg(childrenOf(n, tag));
        if (!f.float) figureHeight += f.heightTwips;
        continue;
      }
      walk(childrenOf(n, tag));
    }
  };
  walk(root);
  return { width, paras, figureHeight };
}

interface SectionSources {
  /** r:id → パートの XML (ヘッダー・フッター) */
  parts: Map<string, XmlNode[]>;
  sizePt: number;
  lineFactor: number;
}

function storyHeight(root: XmlNode[] | undefined, textWidth: number, src: SectionSources): number {
  if (!root) return 0;
  const { width, paras, figureHeight } = storyLength(root);
  if (paras === 0 && figureHeight === 0) return 0;
  const charsPerLine = Math.max(1, textWidth / (src.sizePt * 20));
  const lines = Math.max(paras, Math.ceil(width / charsPerLine));
  return lines * src.sizePt * 20 * src.lineFactor + figureHeight;
}

function parseSectionProps(sectPr: XmlNode | undefined, src: SectionSources, fontSizePt: number): SectionProps {
  const kids = sectPr ? childrenOf(sectPr, tagOf(sectPr)!) : [];
  const pgSz = findChild(kids, "pgSz");
  const pgMar = findChild(kids, "pgMar");
  const docGrid = findChild(kids, "docGrid");
  const cols = findChild(kids, "cols");
  const type = findChild(kids, "type");
  const dir = findChild(kids, "textDirection");
  const pos = (v: number | undefined, fallback: number) => (v !== undefined && v !== 0 ? Math.abs(v) : v === 0 ? 0 : fallback);
  let pageWidth = numAttr(pgSz, "w") ?? 11906; // A4
  let pageHeight = numAttr(pgSz, "h") ?? 16838;
  // 向きの指定と幅・高さが食い違う場合は向きに合わせる
  const orient = pgSz ? attr(pgSz, "orient") : undefined;
  if ((orient === "landscape" && pageWidth < pageHeight) || (orient === "portrait" && pageWidth > pageHeight)) {
    [pageWidth, pageHeight] = [pageHeight, pageWidth];
  }
  const gridTypeRaw = docGrid ? attr(docGrid, "type") : undefined;
  const gridType =
    gridTypeRaw === "lines" || gridTypeRaw === "linesAndChars" || gridTypeRaw === "snapToChars" ? gridTypeRaw : "default";
  const colKids = cols ? childrenOf(cols, "cols") : [];
  const colDefs = colKids
    .filter((c) => tagOf(c) === "col")
    .map((c) => ({ width: numAttr(c, "w") ?? 0, space: numAttr(c, "space") ?? 0 }));
  const num = Math.max(1, numAttr(cols, "num") ?? (colDefs.length || 1));
  const equal = isOn((cols && attr(cols, "equalWidth")) ?? "true");
  const startRaw = type ? attr(type, "val") : undefined;
  const props: SectionProps = {
    pageWidth,
    pageHeight,
    marginTop: pos(numAttr(pgMar, "top"), 1440),
    marginBottom: pos(numAttr(pgMar, "bottom"), 1440),
    marginLeft: pos(numAttr(pgMar, "left"), 1440),
    marginRight: pos(numAttr(pgMar, "right"), 1440),
    gutter: numAttr(pgMar, "gutter") ?? 0,
    headerDistance: numAttr(pgMar, "header") ?? 720,
    footerDistance: numAttr(pgMar, "footer") ?? 720,
    gridType,
    linePitch: numAttr(docGrid, "linePitch") ?? Math.round(fontSizePt * 20 * 1.5),
    charSpace: numAttr(docGrid, "charSpace") ?? 0,
    columns: {
      num,
      space: numAttr(cols, "space") ?? 720,
      separator: isOn(cols && attr(cols, "sep")),
      widths: !equal && colDefs.length === num ? colDefs : undefined,
    },
    start:
      startRaw === "continuous" || startRaw === "evenPage" || startRaw === "oddPage" || startRaw === "nextColumn"
        ? startRaw
        : "nextPage",
    vertical: !!dir && /^(tbRl|tbRlV|tbLrV|rl|rlV)$/.test(attr(dir, "val") ?? ""),
    headerHeight: 0,
    footerHeight: 0,
  };
  // 既定のヘッダー・フッター (w:headerReference w:type="default") の高さ
  const textWidth = pageWidth - props.marginLeft - props.marginRight;
  const ref = (tag: string) => {
    const refs = kids.filter((k) => tagOf(k) === tag);
    const r = refs.find((k) => attr(k, "type") === "default") ?? refs[0];
    return r ? src.parts.get(attr(r, "id") ?? "") : undefined;
  };
  props.headerHeight = storyHeight(ref("headerReference"), textWidth, src);
  props.footerHeight = storyHeight(ref("footerReference"), textWidth, src);
  return props;
}

function parseSection(sectPr: XmlNode | undefined, src: SectionSources, fontSizePt: number, previous?: DocSection): DocSection {
  const section: DocSection = { props: parseSectionProps(sectPr, src, fontSizePt) };
  // ヘッダー・フッターは前のセクションから引き継ぐ
  if (previous && !findChild(sectPr ? childrenOf(sectPr, "sectPr") : [], "headerReference")) {
    section.props.headerHeight = previous.props.headerHeight;
  }
  if (previous && !findChild(sectPr ? childrenOf(sectPr, "sectPr") : [], "footerReference")) {
    section.props.footerHeight = previous.props.footerHeight;
  }
  const change = sectPr ? findChild(childrenOf(sectPr, "sectPr"), "sectPrChange") : undefined;
  const date = change ? attr(change, "date") : undefined;
  const old = change ? findChild(childrenOf(change, "sectPrChange"), "sectPr") : undefined;
  if (old && date) {
    section.before = parseSectionProps(old, src, fontSizePt);
    section.before.headerHeight = section.props.headerHeight;
    section.before.footerHeight = section.props.footerHeight;
    section.changedAt = new Date(date);
  }
  return section;
}

async function partXml(zip: JSZip, name: string): Promise<XmlNode[] | undefined> {
  const f = zip.file(name);
  return f ? (parser.parse(await f.async("string")) as XmlNode[]) : undefined;
}

/** word/_rels/document.xml.rels の Id → パート名 */
async function documentRels(zip: JSZip): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const root = await partXml(zip, "word/_rels/document.xml.rels");
  const rels = root?.find((n) => tagOf(n) === "Relationships");
  for (const r of rels ? childrenOf(rels, "Relationships") : []) {
    const id = attr(r, "Id");
    const target = attr(r, "Target");
    if (id && target) out.set(id, target.startsWith("/") ? target.slice(1) : `word/${target}`);
  }
  return out;
}

/** 脚注・文末脚注 (区切り線などの特別なものは除く) */
async function parseNotes(zip: JSZip, part: string, tag: "footnote" | "endnote"): Promise<Map<string, NoteInfo>> {
  const out = new Map<string, NoteInfo>();
  const root = await partXml(zip, part);
  const top = root?.find((n) => tagOf(n) === `${tag}s`);
  for (const n of top ? childrenOf(top, `${tag}s`) : []) {
    if (tagOf(n) !== tag) continue;
    const type = attr(n, "type");
    if (type && type !== "normal") continue;
    const { width, paras } = storyLength(childrenOf(n, tag));
    out.set(attr(n, "id") ?? "", { width, paras: Math.max(1, paras) });
  }
  return out;
}

export async function parseDocxLayout(input: DocxInput): Promise<DocxLayoutModel> {
  const zip = await JSZip.loadAsync(input);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error(t("errNoDocumentXml"));
  const docRoot = parser.parse(await docFile.async("string")) as XmlNode[];
  const stylesRoot = await partXml(zip, "word/styles.xml");
  const styles = parseStyles(stylesRoot);
  const themeFile = Object.keys(zip.files).find((n) => /^word\/theme\/theme\d*\.xml$/.test(n));
  const theme = parseThemeFonts(themeFile ? await zip.file(themeFile)!.async("string") : undefined);

  // docDefaults
  const defaults = stylesRoot ? findDeep(stylesRoot, "docDefaults") : undefined;
  const defKids = defaults ? childrenOf(defaults, "docDefaults") : [];
  const pPrDefault = findChild(defKids, "pPrDefault");
  const rPrDefault = findChild(defKids, "rPrDefault");
  const sc: StyleContext = {
    styles,
    numbering: parseNumbering(await partXml(zip, "word/numbering.xml")),
    theme,
    defaultPPr: parsePPr(pPrDefault ? findChild(childrenOf(pPrDefault, "pPrDefault"), "pPr") : undefined),
    defaultRPr: parseRPr(rPrDefault ? findChild(childrenOf(rPrDefault, "rPrDefault"), "rPr") : undefined),
    defaultParaStyle: [...styles.entries()].find(([, s]) => s.type === "paragraph" && s.isDefault)?.[0],
  };

  const documentNode = docRoot.find((n) => tagOf(n) === "document");
  const body = documentNode ? findChild(childrenOf(documentNode, "document"), "body") : undefined;
  if (!body) throw new Error(t("errNoBody"));

  const paras: Para[] = [];
  const ctx: BlockContext = { sc, out: paras, tables: [], tableCount: 0, state: {}, section: 0, sectPrs: [] };
  const blocks: BlockRef[] = [];
  const bodySectPr = walkBlocks(childrenOf(body, "body"), undefined, ctx, blocks);

  // 本文の文字サイズ: 既定の段落スタイル > docDefaults > 10.5pt
  const base = resolveParaProps({}, undefined, sc).props;
  const fontSizePt = base.sizePt;

  const rels = await documentRels(zip);
  const parts = new Map<string, XmlNode[]>();
  for (const [id, target] of rels) {
    if (/\/(header|footer)\d*\.xml$/.test(target)) {
      const root = await partXml(zip, target);
      if (root) parts.set(id, root);
    }
  }
  const headerStyle = resolveParaProps({}, [...styles.entries()].find(([, s]) => /^header$/i.test(s.name ?? ""))?.[0], sc).props;
  const src: SectionSources = { parts, sizePt: headerStyle.sizePt, lineFactor: headerStyle.lineFactor };
  const sections: DocSection[] = [];
  for (const sp of [...ctx.sectPrs, bodySectPr]) sections.push(parseSection(sp, src, fontSizePt, sections[sections.length - 1]));

  const appFile = zip.file("docProps/app.xml");
  const pagesMatch = appFile ? (await appFile.async("string")).match(/<Pages>(\d+)<\/Pages>/) : null;
  const savedPageCount = pagesMatch ? Number(pagesMatch[1]) : undefined;

  const settings = await zip.file("word/settings.xml")?.async("string");
  const tabStop = Number(settings?.match(/<w:defaultTabStop\b[^>]*w:val="(\d+)"/)?.[1]);
  const noteStyle = [...styles.entries()].find(([, s]) => /^footnote text$/i.test(s.name ?? ""))?.[0];

  const last = sections[sections.length - 1].props;
  const geometry: PageGeometry = {
    pageWidthTwips: last.pageWidth,
    pageHeightTwips: last.pageHeight,
    marginTopTwips: last.marginTop,
    marginBottomTwips: last.marginBottom,
    marginLeftTwips: last.marginLeft,
    marginRightTwips: last.marginRight,
    fontSizePt,
    lineHeightTwips: last.linePitch,
  };

  return {
    paras,
    geometry,
    savedPageCount: savedPageCount && savedPageCount > 0 ? savedPageCount : undefined,
    sections,
    body: blocks,
    tables: ctx.tables,
    footnotes: await parseNotes(zip, "word/footnotes.xml", "footnote"),
    endnotes: await parseNotes(zip, "word/endnotes.xml", "endnote"),
    noteSizePt: noteStyle ? resolveParaProps({}, noteStyle, sc).props.sizePt : Math.max(8, fontSizePt - 1),
    defaultTabStop: Number.isFinite(tabStop) && tabStop > 0 ? tabStop : 720,
    baseSizePt: fontSizePt,
  };
}
