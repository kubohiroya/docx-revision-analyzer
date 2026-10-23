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
import { XMLParser } from "fast-xml-parser";
import type { DocxInput } from "./input";

type XmlNode = Record<string, unknown>;

/** ツールが書く返信の作成者 (次の収集では、このコメントを記録しない) */
export const RECEIPT_AUTHOR = "Docx Revision Analyzer";
const RECEIPT_INITIALS = "DRA";

/** 引用の前後に残す文字数 */
const CONTEXT_CHARS = 32;

export interface CommentAnchor {
  /** 範囲の始まりの段落の w14:paraId */
  paraId?: string;
  /** 範囲の文字列 (削除された文字を除いた、すべての変更を反映した状態) と、その前後 */
  quote: string;
  prefix: string;
  suffix: string;
  /** 範囲に重なる変更 (種類|作成者|日時|本文のハッシュ値) */
  revisions: string[];
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
function innerText(nodes: XmlNode[], withDeleted = true): string {
  let s = "";
  for (const n of nodes) {
    const tag = tagOf(n);
    if (tag === undefined) continue;
    if (tag === "t" || (withDeleted && tag === "delText")) {
      for (const c of childrenOf(n, tag)) if ("#text" in c) s += String(c["#text"]);
    } else {
      s += innerText(childrenOf(n, tag), withDeleted);
    }
  }
  return s;
}

/** 本文を文書順にたどり、コメントの範囲の文字列・段落・重なる変更を集める */
function collectAnchors(docTree: XmlNode[]): Map<string, CommentAnchor> {
  let text = "";
  let paraId: string | undefined;
  const open = new Map<string, { start: number; paraId?: string; revisions: Set<string> }>();
  const ranges = new Map<string, { start: number; end: number; paraId?: string; revisions: Set<string> }>();
  const points = new Map<string, { at: number; paraId?: string }>();
  const walk = (nodes: XmlNode[]) => {
    for (const n of nodes) {
      const tag = tagOf(n);
      if (tag === undefined) continue;
      const kids = childrenOf(n, tag);
      switch (tag) {
        case "p": {
          const prev = paraId;
          paraId = attr(n, "paraId")?.toUpperCase();
          walk(kids);
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
          if (id !== undefined) open.set(id, { start: text.length, paraId, revisions: new Set() });
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
          if (id !== undefined && !points.has(id)) points.set(id, { at: text.length, paraId });
          break;
        }
        default:
          walk(kids);
      }
    }
  };
  walk(docTree);
  const out = new Map<string, CommentAnchor>();
  const anchor = (start: number, end: number, p: string | undefined, revisions: string[]): CommentAnchor => ({
    paraId: p,
    quote: text.slice(start, end),
    prefix: text.slice(Math.max(0, start - CONTEXT_CHARS), start),
    suffix: text.slice(end, end + CONTEXT_CHARS),
    revisions,
  });
  for (const [id, r] of ranges) out.set(id, anchor(r.start, r.end, r.paraId, [...r.revisions].sort()));
  // 範囲の印が無いコメントは、参照の位置 (文字を持たない範囲) にする
  for (const [id, p] of points) if (!out.has(id)) out.set(id, anchor(p.at, p.at, p.paraId, []));
  return out;
}

/** 文書のコメントを読む (comments.xml の順) */
export async function extractComments(input: DocxInput): Promise<DocxComment[]> {
  const zip = await JSZip.loadAsync(input);
  const tree = await partTree(zip, "word/comments.xml");
  const root = tree?.find((n) => tagOf(n) === "comments");
  if (!root) return [];

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
  const anchors = docTree ? collectAnchors(docTree) : new Map<string, CommentAnchor>();

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
    const text = paras.map((p) => innerText(childrenOf(p, "p"), false)).join("\n");
    comments.push({
      id,
      key: durableId ? `d:${durableId}` : `h:${textHash(`${author ?? ""}\u0000${date ?? ""}\u0000${text}`)}`,
      durableId,
      paraId,
      author,
      initials: attr(n, "initials"),
      date,
      dateUtc: durableId ? utc.get(durableId) : undefined,
      text,
      done: paraId ? ex.get(paraId)?.done ?? false : false,
      anchor: anchors.get(id),
    });
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
    else if (s.comment.text !== c.text) out.push({ at: iso, type: "edited", key: c.key, sha256, comment: c });
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
