/**
 * versions.ts
 *
 * 変更履歴の日時を使って、指定した時刻 T の時点の文書 (版) を復元する。
 *
 *  - 日時が T 以前の挿入 (w:ins / w:moveTo) は受け入れ、T より後の挿入は取り除く
 *  - 日時が T 以前の削除 (w:del / w:moveFrom) は受け入れ、T より後の削除は元に戻す
 *  - 段落記号の挿入・削除も同じ規則で扱う。T の時点で段落記号が無ければ、その段落は次の段落とつながる
 *  - 表の行の挿入・削除 (w:trPr の w:ins / w:del) も同じ規則で扱う
 *  - 書式の変更 (w:rPrChange / w:pPrChange など) は、T より後のものを元の書式に戻す
 *  - 日時の無い変更は、flow.ts と同じく「最初から反映済み」として扱う
 *
 * 復元した版には変更履歴を残さない (T の時点ですべて受け入れた状態)。
 * 復元できるのは変更履歴に残っている範囲だけで、記録がオフの間の編集や、承諾・取り消しで消えた変更は含まれない。
 */

import JSZip from "jszip";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { parseDocxLayout } from "./docxLayout";
import type { DocxInput } from "./input";

type XmlNode = Record<string, unknown>;

const XML_OPTIONS = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  processEntities: true,
  htmlEntities: false,
} as const;

const parser = new XMLParser(XML_OPTIONS);
const builder = new XMLBuilder({ ...XML_OPTIONS, suppressEmptyNode: true });

/** 変更履歴を含みうるパート */
const REVISION_PARTS = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;

function tagOf(node: XmlNode): string | undefined {
  return Object.keys(node).find((k) => k !== ":@" && k !== "#text");
}

/** 名前空間の接頭辞を除いた要素名 */
function localName(tag: string): string {
  return tag.slice(tag.indexOf(":") + 1);
}

function childrenOf(node: XmlNode, tag: string): XmlNode[] {
  const c = node[tag];
  return Array.isArray(c) ? (c as XmlNode[]) : [];
}

function withChildren(node: XmlNode, tag: string, children: XmlNode[]): XmlNode {
  return { ...node, [tag]: children };
}

function findChild(nodes: XmlNode[], name: string): XmlNode | undefined {
  return nodes.find((n) => {
    const tag = tagOf(n);
    return tag !== undefined && localName(tag) === name;
  });
}

function dateOf(node: XmlNode): number | undefined {
  const attrs = node[":@"] as Record<string, unknown> | undefined;
  if (!attrs) return undefined;
  const key = Object.keys(attrs).find((k) => localName(k.slice(2)) === "date");
  const ms = key ? Date.parse(String(attrs[key])) : NaN;
  return Number.isNaN(ms) ? undefined : ms;
}

/** 日時の無い変更は「最初から反映済み」とみなす */
function happenedBy(node: XmlNode, at: number): boolean {
  const d = dateOf(node);
  return d === undefined || d <= at;
}

const INSERTS = new Set(["ins", "moveTo"]);
const DELETES = new Set(["del", "moveFrom"]);
/** 版には不要な、移動範囲などの印 */
const MARKERS = new Set([
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "customXmlInsRangeStart",
  "customXmlInsRangeEnd",
  "customXmlDelRangeStart",
  "customXmlDelRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveToRangeStart",
  "customXmlMoveToRangeEnd",
]);

/**
 * 書式の変更を記録する要素 (親の名前 + "Change")。T より後の変更は、記録されている元の書式に戻す。
 * 元の書式に含まれない子要素 (段落記号の書式、セクションのヘッダー・フッター参照) は今の要素から残す。
 */
const PROPERTY_PARENTS: Record<string, { keepFirst?: string[]; keepLast?: string[] }> = {
  rPr: {},
  pPr: { keepLast: ["rPr", "sectPr"] },
  sectPr: { keepFirst: ["headerReference", "footerReference"] },
  tblPr: {},
  tblPrEx: {},
  trPr: {},
  tcPr: {},
  tblGrid: {},
};

/** 削除が取り消された文章の w:delText / w:delInstrText を、ふつうの w:t / w:instrText に戻す */
function undeleteText(nodes: XmlNode[]): XmlNode[] {
  return nodes.map((n) => {
    const tag = tagOf(n);
    if (tag === undefined) return n;
    const name = localName(tag);
    const kids = undeleteText(childrenOf(n, tag));
    const renamed = name === "delText" ? "t" : name === "delInstrText" ? "instrText" : undefined;
    if (!renamed) return withChildren(n, tag, kids);
    const prefix = tag.slice(0, tag.length - name.length);
    const { [tag]: _old, ...rest } = n;
    return { ...rest, [prefix + renamed]: kids };
  });
}

/** 段落記号・表の行が T の時点で存在するか (w:ins / w:del の印から判断する) */
function presentAt(props: XmlNode[], at: number): boolean {
  const ins = props.find((n) => INSERTS.has(localName(tagOf(n) ?? "")));
  const del = props.find((n) => DELETES.has(localName(tagOf(n) ?? "")));
  return (!ins || happenedBy(ins, at)) && !(del && happenedBy(del, at));
}

/** T の時点で段落記号が無い段落 (次の段落とつながる) に付ける印 */
const JOIN_NEXT = Symbol("joinNext");

function transformProperties(node: XmlNode, tag: string, at: number): XmlNode {
  const name = localName(tag);
  const kids = childrenOf(node, tag);
  const change = findChild(kids, `${name}Change`);
  const rest = kids.filter((k) => k !== change);
  if (!change || happenedBy(change, at)) return withChildren(node, tag, transformChildren(rest, at));
  // 変更前の書式に戻す
  const changeTag = tagOf(change)!;
  const oldProps = childrenOf(change, changeTag).find((k) => tagOf(k) !== undefined);
  const oldKids = oldProps ? childrenOf(oldProps, tagOf(oldProps)!) : [];
  const spec = PROPERTY_PARENTS[name];
  const keep = (names: string[] | undefined) =>
    rest.filter((k) => {
      const t = tagOf(k);
      return t !== undefined && (names ?? []).includes(localName(t));
    });
  return withChildren(node, tag, transformChildren([...keep(spec.keepFirst), ...oldKids, ...keep(spec.keepLast)], at));
}

function transformParagraph(node: XmlNode, tag: string, at: number): XmlNode {
  const kids = childrenOf(node, tag);
  const pPr = findChild(kids, "pPr");
  const rPr = pPr ? findChild(childrenOf(pPr, tagOf(pPr)!), "rPr") : undefined;
  const marks = rPr ? childrenOf(rPr, tagOf(rPr)!) : [];
  const out = withChildren(node, tag, transformChildren(kids, at));
  if (!presentAt(marks, at)) (out as Record<symbol, boolean>)[JOIN_NEXT] = true;
  return out;
}

/**
 * 段落記号の無い段落を次の段落とつなげる。つながった段落は、後ろの段落の段落書式 (段落記号の書式) を持つ。
 * 表やセクションの区切りを越えてはつなげない。
 */
function joinParagraphs(nodes: XmlNode[]): XmlNode[] {
  const out: XmlNode[] = [];
  let pending: { node: XmlNode; tag: string; content: XmlNode[] } | undefined;
  const contentOf = (n: XmlNode, tag: string) =>
    childrenOf(n, tag).filter((k) => {
      const t = tagOf(k);
      return t === undefined || localName(t) !== "pPr";
    });
  const flush = () => {
    if (!pending) return;
    const pPr = findChild(childrenOf(pending.node, pending.tag), "pPr");
    out.push(withChildren(pending.node, pending.tag, pPr ? [pPr, ...pending.content] : pending.content));
    pending = undefined;
  };
  for (const n of nodes) {
    const tag = tagOf(n);
    if (tag === undefined || localName(tag) !== "p") {
      if (tag !== undefined && ["tbl", "sectPr", "sdt", "customXml"].includes(localName(tag))) flush();
      out.push(n);
      continue;
    }
    const content = [...(pending?.content ?? []), ...contentOf(n, tag)];
    pending = undefined;
    if ((n as Record<symbol, boolean>)[JOIN_NEXT]) {
      pending = { node: n, tag, content };
      continue;
    }
    const pPr = findChild(childrenOf(n, tag), "pPr");
    out.push(withChildren(n, tag, pPr ? [pPr, ...content] : content));
  }
  flush();
  return out;
}

/** 要素の子を時刻 at の版に変換する */
function transformChildren(nodes: XmlNode[], at: number): XmlNode[] {
  return joinParagraphs(transformNodes(nodes, at));
}

/** 要素の列を変換する。挿入・削除の要素はほどくか取り除くので、結果は親の子の列に直接入る */
function transformNodes(nodes: XmlNode[], at: number): XmlNode[] {
  const out: XmlNode[] = [];
  for (const n of nodes) {
    const tag = tagOf(n);
    if (tag === undefined) {
      out.push(n);
      continue;
    }
    const name = localName(tag);
    const kids = childrenOf(n, tag);
    if (INSERTS.has(name)) {
      if (happenedBy(n, at)) out.push(...transformNodes(kids, at));
    } else if (DELETES.has(name)) {
      if (!happenedBy(n, at)) out.push(...transformNodes(undeleteText(kids), at));
    } else if (MARKERS.has(name)) {
      // 取り除く
    } else if (name === "p") {
      out.push(transformParagraph(n, tag, at));
    } else if (name === "tr") {
      const trPr = findChild(kids, "trPr");
      if (!trPr || presentAt(childrenOf(trPr, tagOf(trPr)!), at)) out.push(withChildren(n, tag, transformChildren(kids, at)));
    } else if (name === "tbl") {
      const table = withChildren(n, tag, transformChildren(kids, at));
      // 行がすべて無くなった表は取り除く
      if (findChild(childrenOf(table, tag), "tr")) out.push(table);
    } else if (name in PROPERTY_PARENTS) {
      out.push(transformProperties(n, tag, at));
    } else if (name === "cellIns" || name === "cellDel" || name === "cellMerge") {
      // セルの挿入・削除の印は取り除く (セルそのものは残す)
    } else {
      out.push(withChildren(n, tag, transformChildren(kids, at)));
    }
  }
  return out;
}

/** 1つのパートの XML を、時刻 at の版に変換する */
export function versionPartXml(xml: string, at: number): string {
  const tree = parser.parse(xml) as XmlNode[];
  return builder.build(transformChildren(tree, at)) as string;
}

/** 読み込んだ .docx から、複数の時刻の版を作るためのもの (パートの読み込みを1回で済ませる) */
export interface VersionSource {
  /** すべての変更 (挿入・削除・移動・段落記号・表の行・書式) の日時。時刻順 */
  revisionDates: Date[];
  /** 日時の無い変更の数 */
  undatedCount: number;
  /** 時刻 at (ミリ秒) の版の .docx を作る */
  docxAt(at: number): Promise<Uint8Array>;
}

const REVISION_ELEMENTS = new Set([...INSERTS, ...DELETES, "cellIns", "cellDel"]);

/** 変更を記録する要素の日時を集める */
function collectDates(nodes: XmlNode[], dates: number[], undated: { n: number }): void {
  for (const n of nodes) {
    const tag = tagOf(n);
    if (tag === undefined) continue;
    const name = localName(tag);
    if (REVISION_ELEMENTS.has(name) || (name.endsWith("Change") && name.slice(0, -6) in PROPERTY_PARENTS)) {
      const d = dateOf(n);
      if (d === undefined) undated.n++;
      else dates.push(d);
    }
    collectDates(childrenOf(n, tag), dates, undated);
  }
}

export async function openVersionSource(input: DocxInput): Promise<VersionSource> {
  const zip = await JSZip.loadAsync(input);
  const parts = new Map<string, XmlNode[]>();
  for (const name of Object.keys(zip.files)) {
    if (REVISION_PARTS.test(name)) parts.set(name, parser.parse(await zip.file(name)!.async("string")) as XmlNode[]);
  }
  const dates: number[] = [];
  const undated = { n: 0 };
  for (const tree of parts.values()) collectDates(tree, dates, undated);
  return {
    revisionDates: dates.sort((a, b) => a - b).map((d) => new Date(d)),
    undatedCount: undated.n,
    async docxAt(at: number): Promise<Uint8Array> {
      const out = await JSZip.loadAsync(input);
      for (const [name, tree] of parts) out.file(name, builder.build(transformChildren(tree, at)) as string);
      return out.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    },
  };
}

/** 時刻 at の版の .docx を作る */
export async function docxVersionAt(input: DocxInput, at: Date): Promise<Uint8Array> {
  return (await openVersionSource(input)).docxAt(at.getTime());
}

/**
 * 変更履歴を含まない .docx の本文をテキストにする (段落ごとに改行。図は含めない)。
 * docxVersionAt の結果に使う。変更履歴が残っている文書では、すべて受け入れた状態のテキストになる。
 */
export async function docxPlainText(input: DocxInput): Promise<string> {
  const model = await parseDocxLayout(input);
  const lines = model.paras.map((p) =>
    p.segs
      .filter((s) => !s.del)
      .map((s) => (s.kind === "text" ? s.text : s.kind === "figure" ? "" : "\n"))
      .join("")
  );
  return lines.join("\n") + "\n";
}

export type VersionLabel = "initial" | "at" | "every" | "session";

export interface VersionPoint {
  at: Date;
  /** なぜこの時刻の版を作るか (initial: 最初の変更の直前、at: 指定、every: 一定間隔、session: 時間区間の終わり) */
  label: VersionLabel;
}

export interface VersionPointOptions {
  /** 指定した時刻 */
  at?: Date[];
  /** 一定間隔 (ミリ秒)。最初の変更の時刻から数え、その間に変更があった区切りだけを版にする */
  everyMs?: number;
  /** 時間区間 (docx-revision-flow と同じ区切り) の終わりごとに版を作るときの、区切りの閾値 (時間) */
  sessionGapHours?: number;
}

/**
 * 版を作る時刻を決める。dates は変更の日時 (VersionSource.revisionDates)。
 * every / session を指定すると、最初の変更の直前の版 (initial) も加える。同じ時刻は1つにまとめ、時刻順に並べる。
 */
export function versionPoints(dates: Date[], opts: VersionPointOptions): VersionPoint[] {
  const points: VersionPoint[] = (opts.at ?? []).map((at) => ({ at, label: "at" as const }));
  const times = dates.map((d) => d.getTime()).sort((a, b) => a - b);
  if (times.length > 0 && (opts.everyMs !== undefined || opts.sessionGapHours !== undefined)) {
    const first = times[0];
    const last = times[times.length - 1];
    points.push({ at: new Date(first - 1), label: "initial" });
    if (opts.everyMs !== undefined) {
      const step = opts.everyMs;
      let i = 0;
      for (let k = 1; ; k++) {
        const t = Math.min(first + k * step, last);
        // 前の区切りから t までに変更があったときだけ版にする
        let changed = false;
        while (i < times.length && times[i] <= t) {
          changed = true;
          i++;
        }
        if (changed) points.push({ at: new Date(t), label: "every" });
        if (t >= last) break;
      }
    }
    if (opts.sessionGapHours !== undefined) {
      // 直前の変更からの間隔が閾値を超えたら区間を区切る (splitIntoSessions と同じ規則)
      const gapMs = Math.max(0, opts.sessionGapHours) * 3600 * 1000;
      times.forEach((t, j) => {
        if (j === times.length - 1 || times[j + 1] - t > gapMs) points.push({ at: new Date(t), label: "session" });
      });
    }
  }
  const byTime = new Map<number, VersionPoint>();
  for (const p of points) if (!byTime.has(p.at.getTime())) byTime.set(p.at.getTime(), p);
  return [...byTime.values()].sort((a, b) => a.at.getTime() - b.at.getTime());
}
