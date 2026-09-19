/**
 * docxRevisions.ts
 *
 * Word (.docx) の変更履歴 (Track Changes) を解析し、
 * 挿入 (w:ins) / 削除 (w:del) イベントの時系列リストを取り出すライブラリ。
 *
 * 前提:
 *  - 入力ファイルは「変更履歴の記録」を有効にした状態で編集された .docx であること。
 *  - Word は挿入された文字列を <w:ins w:id="..." w:author="..." w:date="...">
 *    要素で、削除された文字列を <w:del ...><w:r><w:delText>...</w:delText></w:r></w:del>
 *    要素でラップして記録する。本モジュールはこれらを再帰的に走査して抽出する。
 *
 * 既知の制約 (README.md にも記載):
 *  - w:moveFrom / w:moveTo (ドラッグ&ドロップ等による同一文書内の移動) は
 *    通常の ins/del とは別要素になる場合があり、本バージョンでは対象外。
 *    Word のバージョン/設定によっては移動も ins/del として記録されることがある。
 *  - w:date のタイムスタンプは秒単位の精度しか持たない。
 *  - 書式のみの変更 (w:pPrChange, w:rPrChange) は文字数に影響しないため対象外。
 *  - フッター/ヘッダー/コメント/脚注中の変更履歴は既定では対象外
 *    (必要であれば extractRevisions の第2引数で対象パートを追加できる)。
 */

import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import * as fs from "fs";

export type RevisionType = "ins" | "del";

export interface RevisionEvent {
  type: RevisionType;
  id: string | number | undefined;
  author: string | undefined;
  date: Date;
  /** このイベントで追加/削除された文字数 */
  chars: number;
  /** 抽出元パート (例: word/document.xml) */
  part: string;
}

export interface DocxRevisionData {
  events: RevisionEvent[];
  /** 現在の文書の「確定後」文字数 (挿入は反映・削除は除いた状態) */
  finalCharCount: number;
  /** 追跡開始前から存在していたと推定される文字数 (finalCharCount - 総挿入 + 総削除) */
  baselineCharCount: number;
  totalInserted: number;
  totalDeleted: number;
}

type XmlNode = Record<string, unknown>;

const DEFAULT_PARTS = ["word/document.xml"];

/** #text を含む子ノード配列から文字列を再帰的に取り出す (t / delText のみ対象) */
function collectText(nodes: XmlNode[] | undefined): number {
  if (!nodes) return 0;
  let total = 0;
  for (const node of nodes) {
    const tagKey = Object.keys(node).find((k) => k !== ":@" && k !== "#text");
    if (tagKey === undefined) continue; // 純粋なテキストノードはここでは扱わない (親で処理)
    const children = node[tagKey] as XmlNode[] | undefined;
    if (tagKey === "t" || tagKey === "delText") {
      // children は [{ "#text": "..." }] の形
      if (Array.isArray(children)) {
        for (const c of children) {
          if (typeof c === "object" && c !== null && "#text" in c) {
            total += String((c as Record<string, unknown>)["#text"]).length;
          }
        }
      }
    } else if (tagKey === "tab" || tagKey === "br" || tagKey === "cr") {
      // タブ/改行も1文字としてカウント (簡易近似)
      total += 1;
    } else {
      // その他の要素は再帰的に潜る (入れ子の ins/del を含む可能性がある)
      total += collectText(children);
    }
  }
  return total;
}

/** 文書全体を走査し、「確定後 (accept 済み)」の可視文字数を数える。del の中身は除外する。 */
function collectFinalText(nodes: XmlNode[] | undefined, insideDel: boolean): number {
  if (!nodes) return 0;
  let total = 0;
  for (const node of nodes) {
    const tagKey = Object.keys(node).find((k) => k !== ":@" && k !== "#text");
    if (tagKey === undefined) continue;
    const children = node[tagKey] as XmlNode[] | undefined;
    if (tagKey === "del") {
      // del の中身 (delText) は確定後は存在しない -> スキップ
      continue;
    } else if (tagKey === "t") {
      if (!insideDel && Array.isArray(children)) {
        for (const c of children) {
          if (typeof c === "object" && c !== null && "#text" in c) {
            total += String((c as Record<string, unknown>)["#text"]).length;
          }
        }
      }
    } else if (tagKey === "tab" || tagKey === "br" || tagKey === "cr") {
      if (!insideDel) total += 1;
    } else {
      total += collectFinalText(children, insideDel);
    }
  }
  return total;
}

function getAttr(node: XmlNode, name: string): string | undefined {
  const attrs = node[":@"] as Record<string, unknown> | undefined;
  if (!attrs) return undefined;
  const v = attrs[`@_${name}`];
  return v === undefined ? undefined : String(v);
}

/** ins/del を再帰的に収集する */
function collectEvents(
  nodes: XmlNode[] | undefined,
  part: string,
  out: RevisionEvent[]
): void {
  if (!nodes) return;
  for (const node of nodes) {
    const tagKey = Object.keys(node).find((k) => k !== ":@" && k !== "#text");
    if (tagKey === undefined) continue;
    const children = node[tagKey] as XmlNode[] | undefined;

    if (tagKey === "ins" || tagKey === "del") {
      const dateStr = getAttr(node, "date");
      const chars = collectText(children);
      // 日付属性が無い場合は事実上解析不能なのでスキップ (警告はライブラリ利用者側で判断)
      if (dateStr) {
        out.push({
          type: tagKey,
          id: getAttr(node, "id"),
          author: getAttr(node, "author"),
          date: new Date(dateStr),
          chars,
          part,
        });
      }
    }
    // ins/del の中に入れ子で ins/del が現れる稀なケースにも対応するため常に再帰する
    collectEvents(children, part, out);
  }
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

async function readPart(zip: JSZip, part: string): Promise<XmlNode[] | undefined> {
  const file = zip.file(part);
  if (!file) return undefined;
  const xml = await file.async("string");
  return parser.parse(xml) as XmlNode[];
}

export interface ExtractOptions {
  /** 解析対象とする docx 内パート (既定: word/document.xml のみ) */
  parts?: string[];
}

export async function extractRevisionsFromBuffer(
  buf: Buffer,
  opts: ExtractOptions = {}
): Promise<DocxRevisionData> {
  const zip = await JSZip.loadAsync(buf);
  const parts = opts.parts ?? DEFAULT_PARTS;

  const events: RevisionEvent[] = [];
  let finalCharCount = 0;

  for (const part of parts) {
    const root = await readPart(zip, part);
    if (!root) continue;
    collectEvents(root, part, events);
    finalCharCount += collectFinalText(root, false);
  }

  // 時系列順に並べる。同一秒の場合は文書中の出現順 (収集順) を維持する (Array.sort は安定ソート)。
  events.sort((a, b) => a.date.getTime() - b.date.getTime());

  const totalInserted = events
    .filter((e) => e.type === "ins")
    .reduce((s, e) => s + e.chars, 0);
  const totalDeleted = events
    .filter((e) => e.type === "del")
    .reduce((s, e) => s + e.chars, 0);

  const baselineCharCount = Math.max(0, finalCharCount - totalInserted + totalDeleted);

  return { events, finalCharCount, baselineCharCount, totalInserted, totalDeleted };
}

export async function extractRevisionsFromFile(
  filePath: string,
  opts: ExtractOptions = {}
): Promise<DocxRevisionData> {
  const buf = await fs.promises.readFile(filePath);
  return extractRevisionsFromBuffer(buf, opts);
}
