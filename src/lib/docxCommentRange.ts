import { XMLParser } from "fast-xml-parser";

export const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Text and UTF-16 offsets from the visible paragraph used in the comparison. */
export interface WordCommentRange {
  paragraphId?: string;
  paragraphText: string;
  start: number;
  end: number;
}

type Node = Record<PropertyKey, unknown>;
const metadata = XMLParser.getMetaDataSymbol() as symbol;
const parser = new XMLParser({ preserveOrder: true, ignoreAttributes: false, trimValues: false,
  parseTagValue: false, captureMetaData: true });

interface TextRun {
  position: number;
  limit: number;
  head: string;
  tail: string;
  textStart: string;
  value: string;
  offset: number;
  trackedFormatting: boolean;
}

function name(node: Node): string | undefined {
  return Object.keys(node).find((key) => key !== ":@" && key !== "#text");
}

function children(node: Node, tag: string): Node[] {
  return node[tag] as Node[] ?? [];
}

function xmlPositionLookup(xml: string): (node: Node) => number | undefined {
  const collapsed = [...xml.matchAll(/\r\n/gu)].map((match, index) => match.index - index);
  return (node) => {
    const position = (node[metadata] as { startIndex?: number } | undefined)?.startIndex;
    if (position === undefined) return;
    // The parser counts each CRLF as one position while the original XML has two.
    let left = 0;
    let right = collapsed.length;
    while (left < right) {
      const middle = Math.floor((left + right) / 2);
      if (collapsed[middle] < position) left = middle + 1;
      else right = middle;
    }
    return position + left;
  };
}

function paragraphId(node: Node): string | undefined {
  const attributes = node[":@"] as Record<string, string> | undefined;
  return Object.entries(attributes ?? {}).find(([key]) => key.endsWith(":paraId"))?.[1]?.toUpperCase();
}

/** The position maps only to ordinary Word text runs that can be split without moving other markup. */
function readTextRun(
  xml: string, run: Node, text: Node, prefix: string, offset: number, location: ReturnType<typeof xmlPositionLookup>,
): TextRun | undefined {
  const position = location(run);
  const textPosition = location(text);
  if (position === undefined || textPosition === undefined) return;
  const opening = /^<[^\s/>]+(?:\s+(?:[^"'<>]|"[^"]*"|'[^']*')*)?>/u.exec(xml.slice(textPosition));
  if (!opening || opening[0].endsWith("/>")) return;
  const textStart = opening[0];
  const valueFrom = textPosition + textStart.length;
  const closeText = `</${prefix}:t>`;
  const closeRun = `</${prefix}:r>`;
  const valueTo = xml.indexOf(closeText, valueFrom);
  if (valueTo < 0 || xml.slice(valueFrom, valueTo).includes("<")) return;
  const runEnd = xml.indexOf(closeRun, valueTo + closeText.length);
  if (runEnd < 0 || xml.slice(valueTo + closeText.length, runEnd).trim()) return;
  const value = children(text, `${prefix}:t`).map((node) => String(node["#text"] ?? "")).join("");
  const head = xml.slice(position, textPosition);
  return { position, limit: runEnd + closeRun.length, head, tail: xml.slice(valueTo, runEnd + closeRun.length),
    textStart, value, offset, trackedFormatting: head.includes(`<${prefix}:rPrChange`) };
}

function textRuns(xml: string, paragraph: Node, prefix: string, location: ReturnType<typeof xmlPositionLookup>): TextRun[] | undefined {
  const w = (tag: string) => `${prefix}:${tag}`;
  const empty = new Set(["pPr", "del", "moveFrom", "bookmarkStart", "bookmarkEnd", "proofErr",
    "commentRangeStart", "commentRangeEnd", "permStart", "permEnd"].map(w));
  const runs: TextRun[] = [];
  let offset = 0;
  for (const node of children(paragraph, w("p"))) {
    const tag = name(node);
    if (!tag || empty.has(tag)) continue;
    if (tag !== w("r")) return;
    const content = children(node, tag).filter((child) => name(child) && name(child) !== w("rPr"));
    if (content.every((child) => [w("commentReference"), w("lastRenderedPageBreak")].includes(name(child) ?? ""))) continue;
    if (content.length !== 1 || name(content[0]) !== w("t")) return;
    const run = readTextRun(xml, node, content[0], prefix, offset, location);
    if (!run) return;
    if (run.value) runs.push(run);
    offset += run.value.length;
  }
  return runs;
}

function piece(xml: string, run: TextRun, start: number, end: number): string {
  if (start === end) return "";
  if (start === 0 && end === run.value.length) return xml.slice(run.position, run.limit);
  if (run.trackedFormatting) throw new Error("The selected text would split a tracked formatting change.");
  const value = run.value.slice(start, end);
  let tag = run.textStart;
  if (/^\s|\s$/u.test(value)) {
    tag = tag.replace(/\s+xml:space\s*=\s*(?:"[^"]*"|'[^']*')/u, "");
    tag = `${tag.slice(0, -1)} xml:space="preserve">`;
  }
  return run.head + tag + value.replace(/&/gu, "&amp;").replace(/</gu, "&lt;").replace(/>/gu, "&gt;") + run.tail;
}

export function insertWordCommentMarks(xml: string, range: WordCommentRange, id: string): string {
  const { paragraphText, start, end } = range;
  const splitsSurrogate = (offset: number) => /[\uDC00-\uDFFF]/u.test(paragraphText.charAt(offset));
  if (!/^\d+$/u.test(id) || !Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end <= start || end > paragraphText.length ||
      splitsSurrogate(start) || splitsSurrogate(end)) throw new Error("The comment range is invalid.");
  const location = xmlPositionLookup(xml);
  const tree = parser.parse(xml, true) as Node[];
  const document = tree.find((node) => name(node)?.endsWith(":document"));
  const tag = document && name(document);
  const prefix = tag?.split(":")[0];
  const attributes = document?.[":@"] as Record<string, string> | undefined;
  if (!document || !tag || !prefix || attributes?.[`@_xmlns:${prefix}`] !== WORD_NS) {
    throw new Error("The Word document format is not supported for comment insertion.");
  }
  const body = children(document, tag).find((node) => name(node) === `${prefix}:body`);
  const idHint = range.paragraphId?.toUpperCase();
  const paragraphs = body ? children(body, `${prefix}:body`) : [];
  const matches = paragraphs.filter((node) => name(node) === `${prefix}:p` && (!idHint || paragraphId(node) === idHint))
    .map((node) => textRuns(xml, node, prefix, location))
    .filter((runs): runs is TextRun[] => runs?.map((run) => run.value).join("") === paragraphText);
  if (matches.length !== 1) throw new Error("The selected paragraph could not be located unambiguously in ordinary Word text.");
  const runs = matches[0];
  const first = runs.find((run) => run.offset <= start && start < run.offset + run.value.length);
  const last = runs.find((run) => run.offset < end && end <= run.offset + run.value.length);
  if (!first || !last) throw new Error("The selected text could not be located.");
  const open = `<${prefix}:commentRangeStart ${prefix}:id="${id}"/>`;
  const close = `<${prefix}:commentRangeEnd ${prefix}:id="${id}"/><${prefix}:r><${prefix}:commentReference ${prefix}:id="${id}"/></${prefix}:r>`;
  const from = start - first.offset;
  const to = end - last.offset;
  const before = xml.slice(0, first.position) + piece(xml, first, 0, from) + open;
  const selected = first === last ? piece(xml, first, from, to) :
    piece(xml, first, from, first.value.length) + xml.slice(first.limit, last.position) + piece(xml, last, 0, to);
  return before + selected + close + piece(xml, last, to, last.value.length) + xml.slice(last.limit);
}
