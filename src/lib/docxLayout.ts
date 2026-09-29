/**
 * docxLayout.ts
 *
 * 本文 (word/document.xml) を「段落」の列に変換し、段落内の各テキスト断片が
 * いつ挿入され (w:ins)、いつ削除されたか (w:del) を保持するモデルを作る。
 * このモデルから、任意の時刻 T における文書の状態 (T 以前の変更は反映し、
 * T より後の変更は元に戻した状態) を復元できる。
 *
 * あわせて、模式的なページ割り付けに使う用紙サイズ・余白・行送り・文字サイズを読み取る。
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
  ins?: RevRef;
  del?: RevRef;
}

export interface BreakSeg {
  kind: "lineBreak" | "pageBreak";
  ins?: RevRef;
  del?: RevRef;
}

export interface FigureSeg {
  kind: "figure";
  /** 図の高さ (twip)。不明な場合は既定値 */
  heightTwips: number;
  ins?: RevRef;
  del?: RevRef;
}

export type Seg = TextSeg | BreakSeg | FigureSeg;

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
  geometry: PageGeometry;
  /** Word が最後に保存したときのページ数 (docProps/app.xml の Pages)。不明なら undefined */
  savedPageCount?: number;
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

/** 全角相当の文字は幅1、それ以外は幅0.5 として表示幅を見積もる */
export function displayWidth(text: string): number {
  let w = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    w += cp >= 0x1100 ? 1 : 0.5;
  }
  return w;
}

/** 文書順に走査しながら、現在どの移動範囲の中にいるかを覚えておく */
interface ParseState {
  moveFromName?: string;
  moveToName?: string;
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

/** 図 (drawing / pict / AlternateContent) の高さを twip で求める */
function figureHeight(nodes: XmlNode[]): number {
  const extent = findDeep(nodes, "extent");
  const cy = extent ? Number(attr(extent, "cy")) : NaN;
  if (Number.isFinite(cy) && cy > 0) return cy / EMU_PER_TWIP;
  return DEFAULT_FIGURE_HEIGHT_TWIPS;
}

interface RevContext {
  ins?: RevRef;
  del?: RevRef;
}

/** 段落の子要素を走査し、テキスト断片等を segs に追加する */
function walkInline(nodes: XmlNode[], ctx: RevContext, segs: Seg[], state: ParseState): void {
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined) continue;
    if (trackMoveRange(tag, node, state)) continue;
    const kids = childrenOf(node, tag);
    switch (tag) {
      case "ins":
        walkInline(kids, { ...ctx, ins: revRef(node, false) }, segs, state);
        break;
      case "moveTo":
        walkInline(kids, { ...ctx, ins: revRef(node, true, state.moveToName) }, segs, state);
        break;
      case "del":
        walkInline(kids, { ...ctx, del: revRef(node, false) }, segs, state);
        break;
      case "moveFrom":
        walkInline(kids, { ...ctx, del: revRef(node, true, state.moveFromName) }, segs, state);
        break;
      case "t":
      case "delText": {
        const text = textOf(kids);
        if (text.length > 0) {
          segs.push({ kind: "text", text, chars: [...text].length, width: displayWidth(text), ...ctx });
        }
        break;
      }
      case "tab":
        segs.push({ kind: "text", text: "\t", chars: 1, width: 2, ...ctx });
        break;
      case "br":
      case "cr":
        segs.push({ kind: attr(node, "type") === "page" ? "pageBreak" : "lineBreak", ...ctx });
        break;
      case "drawing":
      case "pict":
      case "object":
      case "AlternateContent":
        segs.push({ kind: "figure", heightTwips: figureHeight(kids), ...ctx });
        break;
      // 書式・フィールドコード等、本文として表示されないもの
      case "rPr":
      case "pPr":
      case "instrText":
      case "delInstrText":
      case "fldChar":
      case "lastRenderedPageBreak":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "proofErr":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
        break;
      default:
        // r, hyperlink, smartTag, sdt, sdtContent, fldSimple, customXml 等は中へ潜る
        walkInline(kids, ctx, segs, state);
    }
  }
}

interface StyleInfo {
  name?: string;
  outlineLvl?: number;
  basedOn?: string;
  fontSizePt?: number;
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
    const info: StyleInfo = {};
    const name = findChild(kids, "name");
    if (name) info.name = attr(name, "val");
    const basedOn = findChild(kids, "basedOn");
    if (basedOn) info.basedOn = attr(basedOn, "val");
    const pPr = findChild(kids, "pPr");
    const outline = pPr ? findChild(childrenOf(pPr, "pPr"), "outlineLvl") : undefined;
    if (outline) info.outlineLvl = Number(attr(outline, "val"));
    const rPr = findChild(kids, "rPr");
    const sz = rPr ? findChild(childrenOf(rPr, "rPr"), "sz") : undefined;
    if (sz) info.fontSizePt = Number(attr(sz, "val")) / 2;
    map.set(id, info);
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

function sectPrBreaksPage(sectPr: XmlNode): boolean {
  const type = findChild(childrenOf(sectPr, "sectPr"), "type");
  const v = type ? attr(type, "val") : undefined;
  return v !== "continuous";
}

function parseParagraph(
  node: XmlNode,
  index: number,
  tableId: number | undefined,
  styles: Map<string, StyleInfo>,
  state: ParseState
): Para {
  const inTable = tableId !== undefined;
  const kids = childrenOf(node, "p");
  const para: Para = {
    index,
    kind: inTable ? "table" : "body",
    tableId,
    segs: [],
    pageBreakBefore: false,
    sectionBreakAfter: false,
  };
  const pPr = findChild(kids, "pPr");
  if (pPr) {
    const pKids = childrenOf(pPr, "pPr");
    const pStyle = findChild(pKids, "pStyle");
    let level = headingLevelOf(pStyle ? attr(pStyle, "val") : undefined, styles);
    const outline = findChild(pKids, "outlineLvl");
    if (outline) {
      const v = Number(attr(outline, "val"));
      if (v >= 0 && v <= 8) level = v;
    }
    if (level !== undefined && !inTable) {
      para.kind = "heading";
      para.headingLevel = level;
    }
    const pbb = findChild(pKids, "pageBreakBefore");
    if (pbb && !["false", "0", "off"].includes(String(attr(pbb, "val") ?? "true"))) {
      para.pageBreakBefore = true;
    }
    const sectPr = findChild(pKids, "sectPr");
    if (sectPr && sectPrBreaksPage(sectPr)) para.sectionBreakAfter = true;
    // 段落記号の挿入/削除は pPr/rPr の中に記録される
    const rPr = findChild(pKids, "rPr");
    if (rPr) {
      const rKids = childrenOf(rPr, "rPr");
      const ins = findChild(rKids, "ins") ?? findChild(rKids, "moveTo");
      const del = findChild(rKids, "del") ?? findChild(rKids, "moveFrom");
      if (ins) para.markIns = revRef(ins, tagOf(ins) === "moveTo", state.moveToName);
      if (del) para.markDel = revRef(del, tagOf(del) === "moveFrom", state.moveFromName);
    }
  }
  walkInline(kids, {}, para.segs, state);
  return para;
}

interface BlockContext {
  styles: Map<string, StyleInfo>;
  out: Para[];
  tableCount: number;
  state: ParseState;
}

/** 本文ブロック (body, tbl, tc, sdtContent 等) を走査して段落を集める */
function walkBlocks(nodes: XmlNode[], tableId: number | undefined, ctx: BlockContext): XmlNode | undefined {
  let lastSectPr: XmlNode | undefined;
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined) continue;
    if (trackMoveRange(tag, node, ctx.state)) continue;
    if (tag === "p") {
      ctx.out.push(parseParagraph(node, ctx.out.length, tableId, ctx.styles, ctx.state));
    } else if (tag === "tbl") {
      // 入れ子の表は外側の表の一部として扱う
      walkBlocks(childrenOf(node, tag), tableId ?? ctx.tableCount++, ctx);
    } else if (tag === "sectPr") {
      lastSectPr = node;
    } else if (["tr", "tc", "sdt", "sdtContent", "customXml", "ins", "del", "moveTo", "moveFrom"].includes(tag)) {
      // 行単位の挿入/削除 (tr 内の trPr/ins) は段落側の記録で代用する
      walkBlocks(childrenOf(node, tag), tableId, ctx);
    }
  }
  return lastSectPr;
}

function num(node: XmlNode | undefined, name: string, fallback: number): number {
  const v = node ? Number(attr(node, name)) : NaN;
  return Number.isFinite(v) && v > 0 ? v : fallback;
}

function parseGeometry(
  sectPr: XmlNode | undefined,
  styles: Map<string, StyleInfo>,
  stylesRoot: XmlNode[] | undefined
): PageGeometry {
  const kids = sectPr ? childrenOf(sectPr, "sectPr") : [];
  const pgSz = findChild(kids, "pgSz");
  const pgMar = findChild(kids, "pgMar");
  const docGrid = findChild(kids, "docGrid");

  // 本文の文字サイズ: 標準スタイル > docDefaults > 10.5pt
  let fontSizePt = 10.5;
  const defaultsSz = stylesRoot ? findDeep(stylesRoot, "rPrDefault") : undefined;
  const sz = defaultsSz ? findDeep(childrenOf(defaultsSz, "rPrDefault"), "sz") : undefined;
  if (sz) fontSizePt = Number(attr(sz, "val")) / 2 || fontSizePt;
  for (const s of styles.values()) {
    if (s.name?.toLowerCase() === "normal" && s.fontSizePt) fontSizePt = s.fontSizePt;
  }

  const geometry: PageGeometry = {
    pageWidthTwips: num(pgSz, "w", 11906), // A4
    pageHeightTwips: num(pgSz, "h", 16838),
    marginTopTwips: num(pgMar, "top", 1440),
    marginBottomTwips: num(pgMar, "bottom", 1440),
    marginLeftTwips: num(pgMar, "left", 1440),
    marginRightTwips: num(pgMar, "right", 1440),
    fontSizePt,
    lineHeightTwips: num(docGrid, "linePitch", fontSizePt * 20 * 1.5),
  };
  return geometry;
}

export async function parseDocxLayout(buf: Buffer): Promise<DocxLayoutModel> {
  const zip = await JSZip.loadAsync(buf);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error(t("errNoDocumentXml"));
  const docRoot = parser.parse(await docFile.async("string")) as XmlNode[];
  const stylesFile = zip.file("word/styles.xml");
  const stylesRoot = stylesFile
    ? (parser.parse(await stylesFile.async("string")) as XmlNode[])
    : undefined;
  const styles = parseStyles(stylesRoot);

  const documentNode = docRoot.find((n) => tagOf(n) === "document");
  const body = documentNode ? findChild(childrenOf(documentNode, "document"), "body") : undefined;
  if (!body) throw new Error(t("errNoBody"));

  const paras: Para[] = [];
  const sectPr = walkBlocks(childrenOf(body, "body"), undefined, { styles, out: paras, tableCount: 0, state: {} });

  const appFile = zip.file("docProps/app.xml");
  const pagesMatch = appFile ? (await appFile.async("string")).match(/<Pages>(\d+)<\/Pages>/) : null;
  const savedPageCount = pagesMatch ? Number(pagesMatch[1]) : undefined;

  return {
    paras,
    geometry: parseGeometry(sectPr, styles, stylesRoot),
    savedPageCount: savedPageCount && savedPageCount > 0 ? savedPageCount : undefined,
  };
}
