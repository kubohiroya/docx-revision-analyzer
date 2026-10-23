/**
 * PDF の埋め込みフォントの本体を取り除き、表示するときは OS のフォントで代用させる (メインプロセス)。
 *
 * PDF では文字の位置 (改行・改ページ) が1文字ずつ固定されているので、フォントを置き換えても割り付けは変わらない
 * (変わるのは字の形だけ)。日本語の文書では埋め込みフォントが容量の大半を占めるため、版ごとに保管する PDF を軽くできる。
 * 試作の結果 (issue #59) に基づき、フォントの種類ごとに次のようにする。
 *  - CID フォント (/Type0): 本体 (FontFile / FontFile2 / FontFile3) を消す。文字の幅 (W) と ToUnicode は残す
 *  - 1バイトの小分けのフォント (/Type1 / /TrueType。LibreOffice は日本語を 256 字ずつの独自の番号のフォントにする):
 *    pdf.js は本体の無い 1バイトのフォントでは ToUnicode を描画に使わず、番号を欧文の文字として解釈して文字化けする。
 *    そこで ToUnicode から「番号 → 字の名前 uniXXXX」の割り当て (/Encoding の /Differences) を作り直してから本体を消す。
 *    ToUnicode が無ければ取り除かない
 *  - /Type3 (字の形を描く命令) は触らない
 *  - フォント名の接頭辞 (一部だけを埋め込んだ印 "ABCDEF+") を外す (pdf.js は接頭辞があると代わりの書体を探さない)
 *  - 名前から明朝系と分かるフォントに serif の印 (Flags の 2) を立てる (OS に同じフォントが無いとき、明朝体で描かせる)。
 *    1バイトのフォントは、記号のフォントの印 (4) を外して非記号 (32) にする
 * 縦書きの字形 (句読点の位置・長音の向き・小書きの仮名) は OS のフォントでは取り戻せないため、縦書きの部分がある文書には
 * 使わない (呼び出し側で判断する)。
 */
import { inflateSync } from "zlib";
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber, PDFObject, PDFRawStream, PDFRef } from "pdf-lib";

export interface StripResult {
  bytes: Uint8Array;
  /** 本体を取り除いたフォントの数 */
  stripped: number;
  /** そのうち、割り当てを作り直した 1バイトのフォントの数 */
  reencoded: number;
  /** 取り除けなかったフォントの数 (ToUnicode の無い 1バイトのフォント) */
  kept: number;
}

const FONT_FILES = ["FontFile", "FontFile2", "FontFile3"].map((k) => PDFName.of(k));
const SERIF = /serif|mincho|minpro|ming|song|batang|times|century|明朝/i;

function streamBytes(s: PDFRawStream): Uint8Array {
  const filter = s.dict.get(PDFName.of("Filter"))?.toString();
  if (!filter) return s.contents;
  if (filter === "/FlateDecode" || filter === "[ /FlateDecode ]") return inflateSync(s.contents);
  throw new Error(`unsupported filter ${filter}`);
}

/** ToUnicode の CMap から、1バイトの番号 → Unicode の文字列 */
export function parseToUnicode(text: string): Map<number, string> {
  const map = new Map<number, string>();
  const hexToStr = (h: string) => {
    const units: number[] = [];
    for (let i = 0; i < h.length; i += 4) units.push(parseInt(h.slice(i, i + 4).padEnd(4, "0"), 16));
    return String.fromCharCode(...units);
  };
  for (const block of text.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const m of block[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g)) map.set(parseInt(m[1], 16), hexToStr(m[2]));
  }
  for (const block of text.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    for (const m of block[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(<[0-9a-fA-F]+>|\[[^\]]*\])/g)) {
      const lo = parseInt(m[1], 16);
      const hi = parseInt(m[2], 16);
      if (m[3].startsWith("[")) {
        [...m[3].matchAll(/<([0-9a-fA-F]+)>/g)].forEach((x, i) => map.set(lo + i, hexToStr(x[1])));
      } else {
        const base = hexToStr(m[3].slice(1, -1));
        for (let c = lo; c <= hi && c - lo < 256; c++) {
          map.set(c, base.slice(0, -1) + String.fromCharCode(base.charCodeAt(base.length - 1) + (c - lo)));
        }
      }
    }
  }
  return map;
}

/** 文字列の字の名前 (uniXXXX。BMP の外の1文字は uXXXXX) */
export function glyphName(s: string): string {
  const cps = [...s].map((c) => c.codePointAt(0)!);
  if (cps.length === 1 && cps[0] > 0xffff) return `u${cps[0].toString(16).toUpperCase()}`;
  return "uni" + [...s].map((c) => c.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")).join("");
}

/** 埋め込みフォントの本体を取り除いた PDF を作る */
export async function stripEmbeddedFonts(input: Uint8Array): Promise<StripResult> {
  const doc = await PDFDocument.load(input, { updateMetadata: false });
  const ctx = doc.context;
  const lookup = (v: PDFObject | undefined) => (v instanceof PDFRef ? ctx.lookup(v) : v);
  let stripped = 0;
  let reencoded = 0;
  let kept = 0;
  for (const [, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFDict) || obj.get(PDFName.of("Type"))?.toString() !== "/Font") continue;
    const subtype = obj.get(PDFName.of("Subtype"))?.toString();
    if (subtype === "/Type3" || subtype === "/CIDFontType0" || subtype === "/CIDFontType2") continue;
    const descendants = lookup(obj.get(PDFName.of("DescendantFonts")));
    const fontDict = descendants instanceof PDFArray ? lookup(descendants.get(0)) : obj;
    if (!(fontDict instanceof PDFDict)) continue;
    const desc = lookup(fontDict.get(PDFName.of("FontDescriptor")));
    if (!(desc instanceof PDFDict) || !FONT_FILES.some((k) => desc.has(k))) continue;
    const simple = subtype === "/Type1" || subtype === "/TrueType" || subtype === "/MMType1";
    if (simple) {
      const tu = lookup(obj.get(PDFName.of("ToUnicode")));
      let map: Map<number, string> | undefined;
      try {
        map = tu instanceof PDFRawStream ? parseToUnicode(Buffer.from(streamBytes(tu)).toString("latin1")) : undefined;
      } catch {
        map = undefined;
      }
      if (!map || map.size === 0) {
        kept++;
        continue;
      }
      const differences: PDFObject[] = [];
      for (const code of [...map.keys()].sort((a, b) => a - b)) differences.push(PDFNumber.of(code), PDFName.of(glyphName(map.get(code)!)));
      obj.set(PDFName.of("Encoding"), ctx.obj({ Type: "Encoding", Differences: differences }));
      reencoded++;
    }
    for (const k of FONT_FILES) {
      const v = desc.get(k);
      if (!v) continue;
      desc.delete(k);
      if (v instanceof PDFRef) ctx.delete(v);
    }
    stripped++;
    const unprefix = (d: PDFDict) => {
      const n = d.get(PDFName.of("BaseFont"))?.toString();
      if (n) d.set(PDFName.of("BaseFont"), PDFName.of(n.replace(/^\/([A-Z]{6}\+)?/, "")));
    };
    unprefix(obj);
    if (fontDict !== obj) unprefix(fontDict);
    const name = (fontDict.get(PDFName.of("BaseFont"))?.toString() ?? "").slice(1);
    if (name) desc.set(PDFName.of("FontName"), PDFName.of(name));
    let flags = Number(desc.get(PDFName.of("Flags"))?.toString() ?? "0") || 0;
    if (SERIF.test(name) && !/sans|gothic/i.test(name)) flags |= 2;
    if (simple) flags = (flags & ~4) | 32;
    desc.set(PDFName.of("Flags"), PDFNumber.of(flags));
  }
  const bytes = stripped ? await doc.save({ useObjectStreams: true }) : input;
  return { bytes, stripped, reencoded, kept };
}
