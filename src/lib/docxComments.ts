/**
 * docxComments.ts
 *
 * Word のコメントを読み、文書の外の記録 (追記だけの出来事の列) に積むための部品と、
 * 収集・保全したコメントへの返信を文書に書き込む処理。
 *
 * 読むもの:
 *  - word/comments.xml            コメントの本文・作成者・日時 (w:date)
 *  - word/commentsExtended.xml    返信の親子 (w15:paraIdParent) と解決済みか (w15:done)
 *  - word/commentsIds.xml         保存をまたいで変わらない識別子 (w16cid:durableId)
 *  - word/commentsExtensible.xml  UTC の日時 (w16cex:dateUtc)
 *  - word/document.xml            コメントの範囲 (w:commentRangeStart / End、w:commentReference)
 *
 * コメントの w:id は保存のたびに振り直されうるため、記録の鍵には durableId を使い、無ければ作成者・日時・本文の
 * ハッシュ値で代える。範囲は、段落 (w14:paraId)・引用 (範囲の文字列と前後の文字列)・範囲に重なる変更で記録する
 * (後で範囲の文字が書き換わっても、どこを指していたかを追えるように)。
 *
 * 返信の書き込み (addCommentReplies) は、コメントの部分と、本文の中の返信の範囲の印 (親のコメントの印の隣) だけを
 * 文字列として書き足す。本文の文字・変更履歴・既存のコメントの ID と日時には触れない。
 *
 * ファイルシステムは使わない (ブラウザでも動く)。
 */

import JSZip from "jszip";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import type { DocxInput } from "./input";
import { commentLinkUrl, type CommentLink } from "./commentReferences";
import { insertWordCommentMarks, WORD_NS, type WordCommentRange } from "./docxCommentRange";
import { readSavedTextVersion } from "./lineage";

type XmlNode = Record<string, unknown>;

/** ツールが書く返信の作成者 (次の収集では、このコメントを記録しない) */
export const RECEIPT_AUTHOR = "Docx Revision Analyzer";
const RECEIPT_INITIALS = "DRA";

/** 引用の前後に残す文字数 */
const CONTEXT_CHARS = 32;

export interface CommentPassage {
  prefix: string;
  quote: string;
  suffix: string;
}

export interface CommentAnchor extends CommentPassage {
  /** 範囲の始まりの段落の w14:paraId */
  paraId?: string;
  /** 範囲に重なる変更 (種類|作成者|日時|本文のハッシュ値) */
  revisions: string[];
}

export interface CommentLocations {
  incomplete?: boolean;
  passages: Map<string, CommentPassage>;
  find?: (anchor: CommentAnchor) => CommentPassage | undefined;
}

export interface DocxComment {
  /** 文書の中の w:id (保存のたびに変わりうる) */
  id: string;
  /** 記録の鍵: durableId があれば "d:<durableId>"、無ければ "h:<作成者・日時・本文のハッシュ値>" */
  key: string;
  durableId?: string;
  /** コメントの最後の段落の w14:paraId (返信の親子の関係に使う) */
  paraId?: string;
  /** 返信なら、親のコメントの鍵 */
  parentKey?: string;
  author?: string;
  initials?: string;
  /** w:date (書かれたまま) */
  date?: string;
  /** commentsExtensible.xml の UTC の日時 */
  dateUtc?: string;
  text: string;
  links?: CommentLink[];
  /** 解決済み (commentsExtended.xml の w15:done) */
  done: boolean;
  /** 本文の中の範囲 (範囲の印が無いコメントは undefined) */
  anchor?: CommentAnchor;
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

/** 32 ビットの FNV-1a を2つの初期値で計算し、16 桁の 16 進数にする (識別用。暗号用ではない) */
export function textHash(s: string): string {
  const fnv = (seed: number) => {
    let h = seed >>> 0;
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 0x01000193) >>> 0;
    }
    return h.toString(16).padStart(8, "0");
  };
  return fnv(0x811c9dc5) + fnv(0x01000193);
}

/** 変更の鍵 (snapshots.ts で同じ変更とみなす条件と同じ: 種類・作成者・日時・文章) */
export function changeKey(type: string, author: string | undefined, date: string | undefined, text: string): string {
  return `${type}|${author ?? ""}|${date ?? ""}|${textHash(text)}`;
}

async function partTree(zip: JSZip, name: string): Promise<XmlNode[] | undefined> {
  const f = zip.file(name);
  return f ? (parser.parse(await f.async("string")) as XmlNode[]) : undefined;
}

/** 要素の中の文字 (w:t / w:delText) */
function innerText(nodes: XmlNode[], withDeleted = true, withBreaks = false): string {
  let s = "";
  for (const n of nodes) {
    const tag = tagOf(n);
    if (tag === undefined) continue;
    if (tag === "t" || (withDeleted && tag === "delText")) {
      for (const c of childrenOf(n, tag)) if ("#text" in c) s += String(c["#text"]);
    } else if (withBreaks && (tag === "br" || tag === "cr" || tag === "tab")) {
      s += tag === "tab" ? "\t" : "\n";
    } else {
      s += innerText(childrenOf(n, tag), withDeleted, withBreaks);
    }
  }
  return s;
}

/** 本文を文書順にたどり、コメントの範囲の文字列・段落・重なる変更を集める */
function collectAnchors(docTree: XmlNode[], locations?: CommentLocations, references?: Set<string>): Map<string, CommentAnchor> {
  let text = "";
  let paraId: string | undefined;
  const open = new Map<string, { start: number; paraId?: string; revisions: Set<string> }>();
  const ranges = new Map<string, { start: number; end: number; paraId?: string; revisions: Set<string> }>();
  const points = new Map<string, { at: number; paraId?: string }>();
  const ambiguous = new Set<string>();
  const paragraphs: { id: string; start: number; end: number }[] = [];
  const walk = (nodes: XmlNode[]) => {
    for (const n of nodes) {
      const tag = tagOf(n);
      if (tag === undefined) continue;
      if (references && (tag === "commentRangeStart" || tag === "commentRangeEnd" || tag === "commentReference")) {
        const id = attr(n, "id");
        if (id !== undefined) references.add(id);
      }
      const kids = childrenOf(n, tag);
      switch (tag) {
        case "p": {
          const prev = paraId;
          paraId = attr(n, "paraId")?.toUpperCase();
          const start = text.length;
          walk(kids);
          if (locations && paraId) paragraphs.push({ id: paraId, start, end: text.length });
          text += "\n";
          paraId = prev;
          break;
        }
        case "t":
          for (const c of kids) if ("#text" in c) text += String(c["#text"]);
          break;
        case "delText":
        case "instrText":
        case "delInstrText":
          break;
        case "tab":
          text += "\t";
          break;
        case "br":
        case "cr":
          text += "\n";
          break;
        case "ins":
        case "del":
        case "moveTo":
        case "moveFrom": {
          if (open.size) {
            const key = changeKey(tag === "moveTo" ? "ins" : tag === "moveFrom" ? "del" : tag, attr(n, "author"), attr(n, "date"), innerText(kids));
            for (const o of open.values()) o.revisions.add(key);
          }
          walk(kids);
          break;
        }
        case "commentRangeStart": {
          const id = attr(n, "id");
          if (id !== undefined) {
            if (open.has(id) || ranges.has(id)) ambiguous.add(id);
            open.set(id, { start: text.length, paraId, revisions: new Set() });
          }
          break;
        }
        case "commentRangeEnd": {
          const id = attr(n, "id");
          const o = id !== undefined ? open.get(id) : undefined;
          if (id !== undefined && o) {
            ranges.set(id, { start: o.start, end: text.length, paraId: o.paraId, revisions: o.revisions });
            open.delete(id);
          }
          break;
        }
        case "commentReference": {
          const id = attr(n, "id");
          if (id !== undefined) {
            if (points.has(id) && !ranges.has(id)) ambiguous.add(id);
            if (!points.has(id)) points.set(id, { at: text.length, paraId });
          }
          break;
        }
        default:
          walk(kids);
      }
    }
  };
  walk(docTree);
  function passage(start: number, end: number): CommentPassage {
    const lineStart = start === 0 ? 0 : text.lastIndexOf("\n", start - 1) + 1;
    const newline = text.indexOf("\n", end);
    return { prefix: text.slice(lineStart, start), quote: text.slice(start, end), suffix: text.slice(end, newline < 0 ? text.length : newline) };
  }
  if (locations) {
    locations.find = (anchor) => {
      if (!anchor.quote.trim()) return undefined;
      const sameParagraph = paragraphs.filter((p) => p.id === anchor.paraId);
      const candidates: { start: number; end: number; context: boolean }[] = [];
      for (let start = text.indexOf(anchor.quote); start >= 0; start = text.indexOf(anchor.quote, start + 1)) {
        const end = start + anchor.quote.length;
        if (sameParagraph.length && !sameParagraph.some((p) => start >= p.start && start < p.end)) continue;
        const context = Boolean(anchor.prefix || anchor.suffix)
          && text.slice(Math.max(0, start - anchor.prefix.length), start) === anchor.prefix
          && text.slice(end, end + anchor.suffix.length) === anchor.suffix;
        if (sameParagraph.length || context) candidates.push({ start, end, context });
      }
      const contextual = candidates.filter((candidate) => candidate.context);
      const match = contextual.length === 1 ? contextual[0] : candidates.length === 1 ? candidates[0] : undefined;
      return match && passage(match.start, match.end);
    };
  }
  const out = new Map<string, CommentAnchor>();
  const record = (id: string, start: number, end: number, p: string | undefined, revisions: string[]): void => {
    const quote = text.slice(start, end);
    out.set(id, {
      paraId: p,
      quote,
      prefix: text.slice(Math.max(0, start - CONTEXT_CHARS), start),
      suffix: text.slice(end, end + CONTEXT_CHARS),
      revisions,
    });
    if (!ambiguous.has(id) && !open.has(id)) locations?.passages.set(id, passage(start, end));
  };
  for (const [id, r] of ranges) record(id, r.start, r.end, r.paraId, [...r.revisions].sort());
  // 範囲の印が無いコメントは、参照の位置 (文字を持たない範囲) にする
  for (const [id, p] of points) if (!out.has(id)) record(id, p.at, p.at, p.paraId, []);
  return out;
}

/** 文書のコメントを読む (comments.xml の順) */
export async function extractComments(input: DocxInput, locations?: CommentLocations): Promise<DocxComment[]> {
  const zip = await JSZip.loadAsync(input);
  const commentsXml = await zip.file("word/comments.xml")?.async("string");
  if (locations && commentsXml !== undefined && XMLValidator.validate(commentsXml) !== true) {
    locations.incomplete = true;
    return [];
  }
  const tree = commentsXml === undefined ? undefined : parser.parse(commentsXml) as XmlNode[];
  const root = tree?.find((n) => tagOf(n) === "comments");
  if (!root) {
    if (locations) {
      const documentXml = await zip.file("word/document.xml")?.async("string");
      locations.incomplete = commentsXml !== undefined ||
        /<(?:[^<>\s:]+:)?comment(?:Reference|RangeStart|RangeEnd)(?=[\s/>])/.test(documentXml ?? "");
    }
    return [];
  }

  // 解決済み・返信の親子 (paraId → { done, parent paraId })
  const ex = new Map<string, { done: boolean; parent?: string }>();
  const exRoot = (await partTree(zip, "word/commentsExtended.xml"))?.find((n) => tagOf(n) === "commentsEx");
  for (const n of exRoot ? childrenOf(exRoot, "commentsEx") : []) {
    if (tagOf(n) !== "commentEx") continue;
    const p = attr(n, "paraId")?.toUpperCase();
    if (p) ex.set(p, { done: attr(n, "done") === "1", parent: attr(n, "paraIdParent")?.toUpperCase() });
  }
  // paraId → durableId、durableId → UTC の日時
  const durable = new Map<string, string>();
  const idsRoot = (await partTree(zip, "word/commentsIds.xml"))?.find((n) => tagOf(n) === "commentsIds");
  for (const n of idsRoot ? childrenOf(idsRoot, "commentsIds") : []) {
    const p = attr(n, "paraId")?.toUpperCase();
    const d = attr(n, "durableId")?.toUpperCase();
    if (tagOf(n) === "commentId" && p && d) durable.set(p, d);
  }
  const utc = new Map<string, string>();
  const extRoot = (await partTree(zip, "word/commentsExtensible.xml"))?.find((n) => tagOf(n) === "commentsExtensible");
  for (const n of extRoot ? childrenOf(extRoot, "commentsExtensible") : []) {
    const d = attr(n, "durableId")?.toUpperCase();
    const at = attr(n, "dateUtc");
    if (tagOf(n) === "commentExtensible" && d && at) utc.set(d, at);
  }
  const docTree = await partTree(zip, "word/document.xml");
  const references = locations && new Set<string>();
  const anchors = docTree ? collectAnchors(docTree, locations, references) : new Map<string, CommentAnchor>();
  const linksById = new Map<string, string>();
  const rels = (await partTree(zip, "word/_rels/comments.xml.rels"))?.find((n) => tagOf(n) === "Relationships");
  for (const rel of rels ? childrenOf(rels, "Relationships") : []) {
    const id = attr(rel, "Id");
    if (tagOf(rel) !== "Relationship" || !id || attr(rel, "TargetMode") !== "External") continue;
    if (!attr(rel, "Type")?.endsWith("/hyperlink")) continue;
    const url = commentLinkUrl(attr(rel, "Target"));
    if (url) linksById.set(id, url);
  }
  function readLinks(nodes: XmlNode[]): CommentLink[] {
    return nodes.flatMap((node) => {
      const tag = tagOf(node);
      if (!tag) return [];
      const kids = childrenOf(node, tag);
      if (tag !== "hyperlink") return readLinks(kids);
      const url = linksById.get(attr(node, "id") ?? "");
      return url ? [{ url, label: innerText(kids, false) }] : [];
    });
  }

  const comments: DocxComment[] = [];
  for (const n of childrenOf(root, "comments")) {
    if (tagOf(n) !== "comment") continue;
    const kids = childrenOf(n, "comment");
    const paras = kids.filter((k) => tagOf(k) === "p");
    const paraIds = paras.map((p) => attr(p, "paraId")?.toUpperCase()).filter((v): v is string => !!v);
    // commentsExtended / commentsIds はコメントの最後の段落の paraId で引く
    const paraId = [...paraIds].reverse().find((p) => ex.has(p) || durable.has(p)) ?? paraIds[paraIds.length - 1];
    const durableId = paraId ? durable.get(paraId) : undefined;
    const id = attr(n, "id") ?? "";
    const author = attr(n, "author");
    const date = attr(n, "date");
    const text = paras.map((p) => innerText(childrenOf(p, "p"), false, true)).join("\n");
    // Keep the archive keys created before line breaks within runs were extracted.
    const keyText = durableId ? text : paras.map((p) => innerText(childrenOf(p, "p"), false)).join("\n");
    const links = linksById.size ? readLinks(paras) : [];
    comments.push({
      id,
      key: durableId ? `d:${durableId}` : `h:${textHash(`${author ?? ""}\u0000${date ?? ""}\u0000${keyText}`)}`,
      durableId,
      paraId,
      author,
      initials: attr(n, "initials"),
      date,
      dateUtc: durableId ? utc.get(durableId) : undefined,
      text,
      ...(links.length ? { links } : {}),
      done: paraId ? ex.get(paraId)?.done ?? false : false,
      anchor: anchors.get(id),
    });
  }
  if (locations && references) {
    const ids = new Set(comments.map(comment => comment.id));
    locations.incomplete = [...references].some(id => !ids.has(id));
  }
  // 返信の親
  const byPara = new Map(comments.filter((c) => c.paraId).map((c) => [c.paraId!, c]));
  for (const c of comments) {
    const parent = c.paraId ? ex.get(c.paraId)?.parent : undefined;
    if (parent && byPara.has(parent)) c.parentKey = byPara.get(parent)!.key;
  }
  return comments;
}

/** ツールが書いた「保全済み」の返信か */
export function isReceipt(c: DocxComment): boolean {
  return c.author === RECEIPT_AUTHOR;
}

// ---------------------------------------------------------------------------
// 記録 (追記だけの出来事の列)
// ---------------------------------------------------------------------------

/**
 * コメントの出来事:
 *  - seen: 初めて見た (以降、この出来事の seq がコメントの記録番号)
 *  - edited: 本文が変わった / resolved, reopened: 解決済みになった・戻った
 *  - removed: 文書から消えた / reappeared: 消えた後にまた現れた (元に戻した場合など)
 *  - receipt: 「保全済み」の返信を文書に書き込んだ
 */
export type CommentEventType = "seen" | "edited" | "resolved" | "reopened" | "removed" | "reappeared" | "receipt";

export interface CommentEvent {
  /** 記録の通し番号 (1 から) */
  seq: number;
  /** 記録した日時 (ISO 8601) */
  at: string;
  type: CommentEventType;
  key: string;
  /** 出来事を見た版の SHA-256 */
  sha256?: string;
  /** seen / edited / reappeared では、その時点のコメント */
  comment?: DocxComment;
}

export interface CommentState {
  key: string;
  /** 最後に見たときのコメント */
  comment: DocxComment;
  /** 記録番号 (初めて見たときの seq) */
  seq: number;
  firstSeen: string;
  /** いまの文書にあるか */
  present: boolean;
  removedAt?: string;
  /** 「保全済み」の返信を書き込んだ日時 */
  receiptAt?: string;
}

/** 出来事の列から、コメントごとの現在の状態を作る */
export function replayCommentEvents(events: CommentEvent[]): Map<string, CommentState> {
  const state = new Map<string, CommentState>();
  for (const e of events) {
    const s = state.get(e.key);
    switch (e.type) {
      case "seen":
        if (e.comment) state.set(e.key, { key: e.key, comment: e.comment, seq: e.seq, firstSeen: e.at, present: true });
        break;
      case "edited":
      case "reappeared":
        if (s && e.comment) Object.assign(s, { comment: e.comment, present: true, removedAt: undefined });
        break;
      case "resolved":
      case "reopened":
        if (s) s.comment = { ...s.comment, done: e.type === "resolved" };
        break;
      case "removed":
        if (s) Object.assign(s, { present: false, removedAt: e.at });
        break;
      case "receipt":
        if (s) s.receiptAt = e.at;
        break;
    }
  }
  return state;
}

/**
 * いまの文書のコメントを、これまでの状態と比べて出来事にする (seq は記録するときに付ける)。
 * ツールが書いた返信は記録しない
 */
export function diffComments(
  state: Map<string, CommentState>,
  comments: DocxComment[],
  at: Date,
  sha256?: string
): Omit<CommentEvent, "seq">[] {
  const out: Omit<CommentEvent, "seq">[] = [];
  const iso = at.toISOString();
  const seen = new Set<string>();
  for (const c of comments) {
    if (isReceipt(c) || seen.has(c.key)) continue;
    seen.add(c.key);
    const s = state.get(c.key);
    if (!s) {
      out.push({ at: iso, type: "seen", key: c.key, sha256, comment: c });
      continue;
    }
    if (!s.present) out.push({ at: iso, type: "reappeared", key: c.key, sha256, comment: c });
    else if (s.comment.text !== c.text || JSON.stringify(s.comment.links) !== JSON.stringify(c.links)) {
      out.push({ at: iso, type: "edited", key: c.key, sha256, comment: c });
    }
    if (s.comment.done !== c.done) out.push({ at: iso, type: c.done ? "resolved" : "reopened", key: c.key, sha256 });
  }
  for (const s of state.values()) {
    if (s.present && !seen.has(s.key)) out.push({ at: iso, type: "removed", key: s.key, sha256 });
  }
  return out;
}

// ---------------------------------------------------------------------------
// 返信の書き込み
// ---------------------------------------------------------------------------

export interface ReplyRequest {
  /** 返信するコメントの鍵 (DocxComment.key) */
  key: string;
  text: string;
}

export interface ReplyResult {
  /** 返信を書き込んだ文書 */
  bytes: Uint8Array;
  /** 書き込んだ返信の鍵 (親のコメントの鍵) */
  written: string[];
  /** 書き込めなかった返信と理由 (missing: 文書に無い、noAnchor: 本文に範囲の印が無い) */
  skipped: { key: string; reason: "missing" | "noAnchor" }[];
}

const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const REL_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const CT_COMMENTS_EXTENDED = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const NS_MCE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

function escXml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
}

function escRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** 名前空間の URI に結びついた接頭辞 (無ければ undefined) */
function prefixFor(xml: string, uri: string): string | undefined {
  return xml.match(new RegExp(`xmlns:([A-Za-z0-9_]+)="${escRe(uri)}"`))?.[1];
}

/** ルート要素の接頭辞 (<w:document ...> の w) */
function rootPrefix(xml: string, local: string): string {
  return xml.match(new RegExp(`<([A-Za-z0-9_]+):${local}[\\s>]`))?.[1] ?? "w";
}

/** ルート要素に名前空間の宣言を加える */
function declareNs(xml: string, local: string, prefix: string, uri: string): string {
  return xml.replace(new RegExp(`(<[A-Za-z0-9_]+:${local})([\\s>])`), `$1 xmlns:${prefix}="${uri}"$2`);
}

/** 8 桁の 16 進数 (0x7FFFFFFF 未満。paraId / durableId の規則) で、used に無いもの */
function newHexId(used: Set<string>, random: () => number): string {
  for (;;) {
    const v = Math.floor(random() * 0x7ffffffe) + 1;
    const id = v.toString(16).toUpperCase().padStart(8, "0");
    if (!used.has(id)) {
      used.add(id);
      return id;
    }
  }
}

/** Word の w:date の形 (秒まで、UTC) */
function wordDate(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, "Z");
}

export interface NewWordComment {
  range: WordCommentRange;
  author: string;
  text: string;
}

/** Produce a new DOCX in memory, retaining other parts and existing comment threads. */
export async function addWordComment(
  input: DocxInput, comment: NewWordComment, now = new Date(),
): Promise<{ bytes: Uint8Array; id: string }> {
  const author = comment.author.trim();
  const text = comment.text.replace(/\r\n?/gu, "\n");
  const invalidXml = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u;
  if (!text.trim() || !author || author === RECEIPT_AUTHOR || /[\t\r\n]/u.test(author) || invalidXml.test(author + text)) {
    throw new Error("A valid comment and author are required.");
  }
  const zip = await JSZip.loadAsync(input);
  const settings = await partTree(zip, "word/settings.xml");
  const root = settings?.find((node) => tagOf(node) === "settings");
  const restrictions = root && childrenOf(root, "settings").find((node) => tagOf(node) === "documentProtection");
  const enforcement = restrictions && attr(restrictions, "enforcement");
  if (Object.keys(zip.files).some((name) => name.startsWith("_xmlsignatures/")) ||
      (enforcement !== undefined && !["0", "false", "off"].includes(enforcement))) {
    throw new Error("Signed or protected documents cannot be used for comment insertion.");
  }
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) throw new Error("The Word document body is missing.");
  const version = await readSavedTextVersion(input, { label: "", source: "selected" });
  const { paragraphId, paragraphText } = comment.range;
  const matches = !version.unavailable && version.paragraphs.filter((p) =>
    p.text === paragraphText && (!paragraphId || p.paraId?.toUpperCase() === paragraphId.toUpperCase()));
  if (!matches || matches.length !== 1) throw new Error("The selected paragraph is missing or occurs more than once.");

  const part = "word/comments.xml";
  const existing = await zip.file(part)?.async("string");
  await ensureWordCommentsRelationship(zip, existing !== undefined);
  let commentsXml = existing ?? `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${NS_W14}" xmlns:mc="${NS_MCE}" mc:Ignorable="w14"></w:comments>`;
  const prefix = rootPrefix(commentsXml, "comments");
  if (prefixFor(commentsXml, WORD_NS) !== prefix) throw new Error("The Word comments part has an unsupported namespace.");
  const w14 = prefixFor(commentsXml, NS_W14) ?? "w14";
  if (!prefixFor(commentsXml, NS_W14)) commentsXml = declareNs(commentsXml, "comments", w14, NS_W14);

  const comments = await extractComments(input);
  const usedCommentIds = new Set(comments.map(({ id }) => id));
  for (const match of documentXml.matchAll(/<[^:\s]+:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*\b[^:\s]+:id\s*=\s*["'](\d+)["']/gu)) {
    usedCommentIds.add(match[1]);
  }
  let id = 0;
  while (usedCommentIds.has(String(id))) id++;
  const key = String(id);
  const marked = insertWordCommentMarks(documentXml, comment.range, key);
  const extendedParts = await Promise.all(["commentsExtended", "commentsIds", "commentsExtensible"].map(async (name) => ({
    name: `word/${name}.xml`, xml: await zip.file(`word/${name}.xml`)?.async("string"),
  })));
  const usedIds = new Set([...`${documentXml}${commentsXml}${extendedParts.map(({ xml }) => xml ?? "").join("")}`
    .matchAll(/(?:paraId|durableId)\s*=\s*["']([0-9A-Fa-f]{8})["']/gu)].map((match) => match[1].toUpperCase()));
  const paraId = newHexId(usedIds, Math.random);
  const durableId = newHexId(usedIds, Math.random);
  const date = wordDate(now);
  const w = (name: string) => `${prefix}:${name}`;
  const body = text.split("\n").map((line) => `<${w("t")} xml:space="preserve">${escXml(line)}</${w("t")}>`).join(`<${w("br")}/>`);
  const node = `<${w("comment")} ${w("id")}="${key}" ${w("author")}="${escXml(author)}" ${w("date")}="${date}">` +
    `<${w("p")} ${w14}:paraId="${paraId}"><${w("r")}><${w("annotationRef")}/></${w("r")}>` +
    `<${w("r")}>${body}</${w("r")}></${w("p")}></${w("comment")}>`;
  zip.file(part, insertIntoXmlRoot(commentsXml, w("comments"), node));
  zip.file("word/document.xml", marked);
  const additions = [
    ["commentsEx", NS_W15, "commentEx", `paraId="${paraId}" done="0"`],
    ["commentsIds", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "commentId", `paraId="${paraId}" durableId="${durableId}"`],
    ["commentsExtensible", "http://schemas.microsoft.com/office/word/2018/wordml/cex", "commentExtensible", `durableId="${durableId}" dateUtc="${date}"`],
  ];
  for (const [index, { name, xml }] of extendedParts.entries()) {
    if (!xml) continue;
    const [root, ns, item, attributes] = additions[index];
    const p = prefixFor(xml, ns);
    if (!p) throw new Error("The Word comment extensions have an unsupported namespace.");
    const attrs = attributes.replace(/\b([A-Za-z]+)=/gu, `${p}:$1=`);
    zip.file(name, insertIntoXmlRoot(xml, `${p}:${root}`, `<${p}:${item} ${attrs}/>`));
  }
  return { bytes: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), id: key };
}

function insertIntoXmlRoot(xml: string, name: string, child: string): string {
  const closing = new RegExp(`</${escRe(name)}\\s*>\\s*$`, "u");
  if (closing.test(xml)) return xml.replace(closing, (end) => child + end);
  const empty = new RegExp(`(<${escRe(name)}(?:\\s[^<>]*?)?)/>(\\s*)$`, "u");
  if (empty.test(xml)) return xml.replace(empty, (_match, start, space) => `${start}>${child}</${name}>${space}`);
  throw new Error(`The Word package has an unsupported ${name} part.`);
}

async function ensureWordCommentsRelationship(zip: JSZip, exists: boolean): Promise<void> {
  const relName = "word/_rels/document.xml.rels";
  const relsXml = await zip.file(relName)?.async("string") ??
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const root = (parser.parse(relsXml) as XmlNode[]).find((node) => tagOf(node) === "Relationships");
  if (!root) throw new Error("The Word document relationships part is invalid.");
  const rels = childrenOf(root, "Relationships");
  const comments = rels.filter((node) => attr(node, "Type") === REL_COMMENTS);
  const regularTarget = (node: XmlNode) => ["comments.xml", "/word/comments.xml"].includes(attr(node, "Target") ?? "") &&
    attr(node, "TargetMode") !== "External";
  if (exists ? comments.length !== 1 || !regularTarget(comments[0])
    : comments.length > 0 || rels.some(regularTarget)) throw new Error("The Word comments part does not have a supported relationship.");
  const typesName = "[Content_Types].xml";
  const types = await zip.file(typesName)?.async("string");
  if (!types) throw new Error("The Word package content types are missing.");
  const typeRoot = (parser.parse(types) as XmlNode[]).find((node) => tagOf(node) === "Types");
  const overrides = typeRoot && childrenOf(typeRoot, "Types").filter((node) => attr(node, "PartName") === "/word/comments.xml");
  if (!overrides || (exists ? overrides.length !== 1 || attr(overrides[0], "ContentType") !== CT_COMMENTS : overrides.length !== 0)) {
    throw new Error("The Word package content types are invalid.");
  }
  if (exists) return;
  const ids = new Set(rels.map((node) => attr(node, "Id")));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  zip.file(relName, insertIntoXmlRoot(relsXml, "Relationships", `<Relationship Id="rId${n}" Type="${REL_COMMENTS}" Target="comments.xml"/>`));
  zip.file(typesName, insertIntoXmlRoot(types, "Types", `<Override PartName="/word/comments.xml" ContentType="${CT_COMMENTS}"/>`));
}

/**
 * コメントに返信を書き込む。返信は RECEIPT_AUTHOR の名前で、親のコメントと同じ範囲に付ける。
 * 書き足すのはコメントの部分 (comments.xml、commentsExtended.xml、あれば commentsIds.xml・commentsExtensible.xml) と、
 * 本文の中の親のコメントの印の隣に置く返信の印だけ
 */
export async function addCommentReplies(
  input: DocxInput,
  replies: ReplyRequest[],
  opts: { author?: string; initials?: string; now?: Date; random?: () => number } = {}
): Promise<ReplyResult> {
  const zip = await JSZip.loadAsync(input);
  const comments = await extractComments(input);
  const byKey = new Map(comments.map((c) => [c.key, c]));
  const random = opts.random ?? Math.random;
  const now = opts.now ?? new Date();
  const author = escXml(opts.author ?? RECEIPT_AUTHOR);
  const initials = escXml(opts.initials ?? RECEIPT_INITIALS);

  let commentsXml = await zip.file("word/comments.xml")?.async("string");
  let documentXml: string | undefined = await zip.file("word/document.xml")?.async("string");
  const result: ReplyResult = { bytes: new Uint8Array(), written: [], skipped: [] };
  if (!commentsXml || !documentXml) {
    result.skipped = replies.map((r) => ({ key: r.key, reason: "missing" as const }));
    result.bytes = input instanceof Uint8Array ? input : new Uint8Array(input as ArrayBuffer);
    return result;
  }
  let extendedXml = await zip.file("word/commentsExtended.xml")?.async("string");
  let idsXml = await zip.file("word/commentsIds.xml")?.async("string");
  let extensibleXml = await zip.file("word/commentsExtensible.xml")?.async("string");

  const cw = rootPrefix(commentsXml, "comments");
  const dw = rootPrefix(documentXml, "document");
  let w14 = prefixFor(commentsXml, NS_W14);
  if (!w14) {
    w14 = "w14";
    commentsXml = declareNs(commentsXml, "comments", w14, NS_W14);
  }
  const used = new Set(
    [...`${commentsXml}${documentXml}${extendedXml ?? ""}${idsXml ?? ""}`.matchAll(/(?:paraId|durableId|textId)="([0-9A-Fa-f]{8})"/g)].map((m) =>
      m[1].toUpperCase()
    )
  );
  let nextId = Math.max(0, ...[...commentsXml.matchAll(new RegExp(`<${cw}:comment\\b[^>]*\\b${cw}:id="(\\d+)"`, "g"))].map((m) => Number(m[1]))) + 1;

  const newComments: string[] = [];
  const newEx: string[] = [];
  const newIds: string[] = [];
  const newExtensible: string[] = [];
  for (const r of replies) {
    const parent = byKey.get(r.key);
    if (!parent) {
      result.skipped.push({ key: r.key, reason: "missing" });
      continue;
    }
    const pid = escRe(parent.id);
    const startRe = new RegExp(`<${dw}:commentRangeStart\\b[^>]*\\b${dw}:id="${pid}"[^>]*/>`);
    const endRe = new RegExp(`<${dw}:commentRangeEnd\\b[^>]*\\b${dw}:id="${pid}"[^>]*/>`);
    const refRe = new RegExp(`<${dw}:commentReference\\b[^>]*\\b${dw}:id="${pid}"[^>]*/>`);
    const ref = refRe.exec(documentXml);
    const runEnd: number = ref ? documentXml.indexOf(`</${dw}:r>`, ref.index) : -1;
    if (!ref || runEnd < 0) {
      result.skipped.push({ key: r.key, reason: "noAnchor" });
      continue;
    }
    const id = String(nextId++);
    // 本文: 親の範囲の印の隣に返信の範囲の印を、親の参照の後に返信の参照を置く (後ろから書き足して位置をずらさない)
    const refRun = `<${dw}:r><${dw}:commentReference ${dw}:id="${id}"/></${dw}:r>`;
    documentXml = documentXml.slice(0, runEnd + `</${dw}:r>`.length) + refRun + documentXml.slice(runEnd + `</${dw}:r>`.length);
    documentXml = documentXml.replace(endRe, (m) => `${m}<${dw}:commentRangeEnd ${dw}:id="${id}"/>`);
    documentXml = documentXml.replace(startRe, (m) => `${m}<${dw}:commentRangeStart ${dw}:id="${id}"/>`);

    const paraId = newHexId(used, random);
    const durableId = newHexId(used, random);
    newComments.push(
      `<${cw}:comment ${cw}:id="${id}" ${cw}:author="${author}" ${cw}:date="${wordDate(now)}" ${cw}:initials="${initials}">` +
        `<${cw}:p ${w14}:paraId="${paraId}" ${w14}:textId="77777777"><${cw}:r><${cw}:annotationRef/></${cw}:r>` +
        `<${cw}:r><${cw}:t xml:space="preserve">${escXml(r.text)}</${cw}:t></${cw}:r></${cw}:p></${cw}:comment>`
    );
    newEx.push(paraId + (parent.paraId ? `|${parent.paraId}` : ""));
    newIds.push(`${paraId}|${durableId}`);
    newExtensible.push(durableId);
    result.written.push(r.key);
  }
  if (result.written.length === 0) {
    result.bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
    return result;
  }

  commentsXml = commentsXml.replace(new RegExp(`</${cw}:comments>\\s*$`), `${newComments.join("")}</${cw}:comments>`);
  zip.file("word/comments.xml", commentsXml);
  zip.file("word/document.xml", documentXml);

  // 返信の親子 (commentsExtended.xml)。無ければ作る
  if (!extendedXml) {
    extendedXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="${NS_W15}" mc:Ignorable="w15"></w15:commentsEx>`;
    await addPart(zip, "commentsExtended.xml", REL_COMMENTS_EXTENDED, CT_COMMENTS_EXTENDED);
  }
  const w15 = prefixFor(extendedXml, NS_W15) ?? "w15";
  const exItems = newEx.map((v) => {
    const [p, parentPara] = v.split("|");
    return `<${w15}:commentEx ${w15}:paraId="${p}"${parentPara ? ` ${w15}:paraIdParent="${parentPara}"` : ""} ${w15}:done="0"/>`;
  });
  extendedXml = extendedXml.replace(new RegExp(`</${w15}:commentsEx>\\s*$`), `${exItems.join("")}</${w15}:commentsEx>`);
  zip.file("word/commentsExtended.xml", extendedXml);

  // 識別子と UTC の日時 (Word 2019 以降の部分。文書にあるときだけ書き足す)
  if (idsXml) {
    const p = prefixFor(idsXml, "http://schemas.microsoft.com/office/word/2016/wordml/cid") ?? "w16cid";
    const items = newIds.map((v) => {
      const [para, dur] = v.split("|");
      return `<${p}:commentId ${p}:paraId="${para}" ${p}:durableId="${dur}"/>`;
    });
    idsXml = idsXml.replace(new RegExp(`</${p}:commentsIds>\\s*$`), `${items.join("")}</${p}:commentsIds>`);
    zip.file("word/commentsIds.xml", idsXml);
  }
  if (extensibleXml) {
    const p = prefixFor(extensibleXml, "http://schemas.microsoft.com/office/word/2018/wordml/cex") ?? "w16cex";
    const items = newExtensible.map((dur) => `<${p}:commentExtensible ${p}:durableId="${dur}" ${p}:dateUtc="${wordDate(now)}"/>`);
    extensibleXml = extensibleXml.replace(new RegExp(`</${p}:commentsExtensible>\\s*$`), `${items.join("")}</${p}:commentsExtensible>`);
    zip.file("word/commentsExtensible.xml", extensibleXml);
  }
  result.bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return result;
}

/** word/ の下にパートを加える (document.xml からの関係と、[Content_Types].xml の型) */
async function addPart(zip: JSZip, name: string, relType: string, contentType: string): Promise<void> {
  const relsName = "word/_rels/document.xml.rels";
  let rels = await zip.file(relsName)?.async("string");
  if (rels && !rels.includes(`Target="${name}"`)) {
    const ids = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
    let n = ids.size + 1;
    while (ids.has(`rId${n}`)) n++;
    rels = rels.replace(/<\/Relationships>\s*$/, `<Relationship Id="rId${n}" Type="${relType}" Target="${name}"/></Relationships>`);
    zip.file(relsName, rels);
  }
  let types = await zip.file("[Content_Types].xml")?.async("string");
  if (types && !types.includes(`PartName="/word/${name}"`)) {
    types = types.replace(/<\/Types>\s*$/, `<Override PartName="/word/${name}" ContentType="${contentType}"/></Types>`);
    zip.file("[Content_Types].xml", types);
  }
}
