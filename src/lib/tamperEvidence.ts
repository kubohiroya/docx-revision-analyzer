/**
 * tamperEvidence.ts
 *
 * 変更履歴の「改ざんの痕跡」を調べる。integrity.ts (情報として示すだけ) と違い、ここで見つかった痕跡は
 * 図の警告と出力ファイル名 (-tampered.svg) に使う。
 *
 * 想定する使い方: 教員が「変更履歴のロック」をかけたテンプレートを配り、学生が書いて提出したファイルを
 * 調べる。ロックは settings.xml を書き換えれば外せ、本文を別の文書に写せば変更履歴ごと消えるため、
 * 提出時に次の痕跡を確かめる。
 *
 * テンプレートが無くても調べる項目:
 *  - trackingOff:       「変更履歴の記録」がオフの状態で保存されている
 *  - lockReleased:      変更履歴のロック (w:documentProtection w:edit="trackedChanges") が解除されている
 *                        (テンプレートがあれば、テンプレートでロックされていたのに提出物でロックされていない場合も)
 *  - authorDateRemoved: 作成者や日時の無い変更がある (時系列が失われている)
 * テンプレート (--template) を渡したときだけ調べる項目:
 *  - notFromTemplate:   テンプレートから作られた文書ではない (最初の編集セッションの識別子 rsidRoot が違う)
 *  - lockChanged:       ロックのパスワード (ハッシュ値) がテンプレートと違う
 *  - untrackedText:     テンプレートに無い編集セッション (rsid) で入力された本文が、変更履歴 (w:ins) の外にある。
 *                        記録をオフにして入力したか、変更を承諾したときに残る
 * スナップショットの通し解析 (snapshots.ts) で加える項目:
 *  - datedBeforePrevious: 前回の提出より前の日時の変更が、今回の提出で初めて現れた
 *
 * rsid の照合をテンプレート無しで行わないのは、文書を作って少し書いてから同じセッションのうちに記録をオンにする、
 * という正当な使い方と区別できないため。
 *
 * どの痕跡も、改ざんの証拠ではない (Word 以外のアプリ・古い Word・変換ツールでも同じ形になることがある)。
 */

import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type { DocxInput } from "./input";
import type { LocalizedText } from "./categories";
import { getLang, Lang } from "./i18n";
import { parseHistorySettings } from "./historySettings";

export type TamperEvidenceId =
  | "trackingOff"
  | "lockReleased"
  | "lockChanged"
  | "authorDateRemoved"
  | "notFromTemplate"
  | "untrackedText"
  /** スナップショットの通し解析 (snapshots.ts) で使う: 前回の提出より前の日時の変更が、今回初めて現れた */
  | "datedBeforePrevious";

export interface TamperEvidence {
  id: TamperEvidenceId;
  message: LocalizedText;
  details: Record<string, string | number | boolean | null>;
}

export interface TamperReport {
  /** テンプレートと照合したか */
  checkedAgainstTemplate: boolean;
  /** 見つかった痕跡 (無ければ空) */
  evidence: TamperEvidence[];
  note: LocalizedText;
}

export const TAMPER_NOTE: LocalizedText = {
  en:
    "These traces are not proof of tampering: apps other than Word, old Word versions and converters can leave the " +
    "same marks. Check the file and ask the writer before drawing conclusions.",
  ja:
    "これらの痕跡は改ざんの証拠ではありません。Word 以外のアプリ・古い Word・変換ツールでも同じ形になることがあります。" +
    "結論を出す前に、ファイルを確かめ、書いた本人に確認してください。",
};

export interface TamperCheckOptions {
  /** 配布したテンプレートの docx。渡すと、テンプレートとの照合も行う */
  template?: DocxInput;
  /** 記録外で入力された本文が、空白を除いて何文字以上あれば痕跡とするか (既定: 1) */
  minUntrackedChars?: number;
  /**
   * ロックのパスワード (ハッシュ値) をテンプレートと比べるか (既定: true)。前回の提出物を基準にする場合、
   * 教員が承諾のためにロックを外してかけ直すと salt が変わるため、false にする
   */
  compareLockPassword?: boolean;
}

type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

function tagOf(n: XmlNode): string | undefined {
  return Object.keys(n).find((k) => k !== ":@" && k !== "#text");
}

function attrs(n: XmlNode): Record<string, string> {
  const a = (n[":@"] as Record<string, unknown> | undefined) ?? {};
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(a)) out[k.replace(/^@_/, "")] = String(v);
  return out;
}

/** 本文の連続 (w:r) と、それが変更履歴の挿入 (w:ins / w:moveTo) の中にあるか */
interface RunInfo {
  rsidR?: string;
  tracked: boolean;
  /** 空白を除いた文字数 */
  chars: number;
}

function runChars(nodes: XmlNode[]): number {
  let n = 0;
  for (const node of nodes) {
    const tag = tagOf(node);
    if (!tag) continue;
    const kids = node[tag];
    if ((tag === "t" || tag === "delText") && Array.isArray(kids)) {
      for (const c of kids as XmlNode[]) if ("#text" in c) n += String(c["#text"]).replace(/\s/g, "").length;
    }
  }
  return n;
}

/** 文書順に本文の連続と、使われている rsid を集める */
function walk(nodes: XmlNode[], tracked: boolean, runs: RunInfo[], rsids: Set<string>): void {
  for (const n of nodes) {
    const tag = tagOf(n);
    if (!tag) continue;
    const a = attrs(n);
    for (const [k, v] of Object.entries(a)) if (/^rsid/.test(k) && /^[0-9A-Fa-f]{8}$/.test(v)) rsids.add(v.toUpperCase());
    const kids = Array.isArray(n[tag]) ? (n[tag] as XmlNode[]) : [];
    if (tag === "r") {
      runs.push({ rsidR: a.rsidR?.toUpperCase(), tracked, chars: runChars(kids) });
      continue;
    }
    // w:ins / w:moveTo は段落記号の変更 (rPr の中) にも現れるが、そこに w:r は無い
    walk(kids, tracked || ((tag === "ins" || tag === "moveTo") && a.id !== undefined), runs, rsids);
  }
}

interface Protection {
  edit?: string;
  enforced: boolean;
  /** パスワードのハッシュ値と salt (形式の違い w:hash / w:hashValue を吸収して連結したもの) */
  key: string;
}

interface DocxFacts {
  settingsPartFound: boolean;
  trackRevisions: boolean;
  protection?: Protection;
  rsidRoot?: string;
  /** settings.xml の rsid の一覧 */
  listedRsids: Set<string>;
  /** 本文で使われている rsid */
  usedRsids: Set<string>;
  runs: RunInfo[];
  undated: number;
  withoutAuthor: number;
}

function readProtection(settingsXml: string | undefined): Protection | undefined {
  const m = settingsXml?.match(/<(?:\w+:)?documentProtection\b([^>]*)\/?>/);
  if (!m) return undefined;
  const a: Record<string, string> = {};
  for (const x of m[1].matchAll(/(?:\w+:)?(\w+)="([^"]*)"/g)) a[x[1]] = x[2];
  const enforcement = (a.enforcement ?? "").toLowerCase();
  return {
    edit: a.edit,
    enforced: ["1", "true", "on"].includes(enforcement),
    key: [a.hashValue ?? a.hash ?? "", a.saltValue ?? a.salt ?? ""].join("|"),
  };
}

function isTrackLock(p: Protection | undefined): boolean {
  return !!p && p.edit === "trackedChanges" && p.enforced;
}

async function readFacts(input: DocxInput): Promise<DocxFacts> {
  const zip = await JSZip.loadAsync(input);
  const read = async (p: string) => (zip.file(p) ? await zip.file(p)!.async("string") : undefined);
  const [docXml, settingsXml] = await Promise.all([read("word/document.xml"), read("word/settings.xml")]);
  const settings = parseHistorySettings(settingsXml);
  const runs: RunInfo[] = [];
  const usedRsids = new Set<string>();
  if (docXml) walk(parser.parse(docXml) as XmlNode[], false, runs, usedRsids);
  const listedRsids = new Set(
    [...(settingsXml?.matchAll(/<(?:\w+:)?rsid\b[^>]*?(?:\w+:)?val="([0-9A-Fa-f]{8})"/g) ?? [])].map((m) => m[1].toUpperCase())
  );
  const rsidRoot = settingsXml?.match(/<(?:\w+:)?rsidRoot\b[^>]*?(?:\w+:)?val="([0-9A-Fa-f]{8})"/)?.[1]?.toUpperCase();
  // 変更 (挿入・削除・移動) の日時と作成者
  let undated = 0;
  let withoutAuthor = 0;
  for (const m of docXml?.matchAll(/<(?:\w+:)?(?:ins|del|moveFrom|moveTo)\b([^>]*)>/g) ?? []) {
    if (!/\s(?:\w+:)?id="/.test(m[1])) continue;
    if (!/\s(?:\w+:)?date="/.test(m[1])) undated++;
    if (!/\s(?:\w+:)?author="/.test(m[1])) withoutAuthor++;
  }
  return {
    settingsPartFound: settings.settingsPartFound,
    trackRevisions: settings.trackRevisions,
    protection: readProtection(settingsXml),
    rsidRoot,
    listedRsids,
    usedRsids,
    runs,
    undated,
    withoutAuthor,
  };
}

/** docx のバイト列から、改ざんの痕跡を調べる */
export async function checkTamperEvidence(input: DocxInput, opts: TamperCheckOptions = {}): Promise<TamperReport> {
  const doc = await readFacts(input);
  const tpl = opts.template ? await readFacts(opts.template) : undefined;
  const evidence: TamperEvidence[] = [];

  // trackingOff
  if (doc.settingsPartFound && !doc.trackRevisions) {
    evidence.push({
      id: "trackingOff",
      details: {},
      message: {
        en: "Track Changes was off when the file was saved. Edits made while it was off are not recorded.",
        ja: "「変更履歴の記録」がオフの状態で保存されています。オフの間の編集は記録されていません。",
      },
    });
  }

  // lockReleased / lockChanged
  const p = doc.protection;
  if (tpl && isTrackLock(tpl.protection) && !isTrackLock(p)) {
    evidence.push({
      id: "lockReleased",
      details: { protection: p ? `${p.edit ?? ""}${p.enforced ? "" : " (not enforced)"}` : null },
      message: {
        en: "The template locked Track Changes, but this file is not locked: the lock was removed.",
        ja: "テンプレートでは変更履歴がロックされていましたが、このファイルはロックされていません。ロックが外されています。",
      },
    });
  } else if (!tpl && p && p.edit === "trackedChanges" && !p.enforced) {
    evidence.push({
      id: "lockReleased",
      details: { protection: "trackedChanges (not enforced)" },
      message: {
        en: "Track Changes was locked once, but the lock has been released.",
        ja: "変更履歴はロックされていたことがありますが、そのロックが解除されています。",
      },
    });
  }
  if (opts.compareLockPassword !== false && tpl && isTrackLock(tpl.protection) && isTrackLock(p) && p!.key !== tpl.protection!.key) {
    evidence.push({
      id: "lockChanged",
      details: {},
      message: {
        en: "Track Changes is locked, but with a different password from the template's.",
        ja: "変更履歴はロックされていますが、パスワードがテンプレートと違います。",
      },
    });
  }

  // authorDateRemoved
  if (doc.undated > 0 || doc.withoutAuthor > 0) {
    evidence.push({
      id: "authorDateRemoved",
      details: { undated: doc.undated, withoutAuthor: doc.withoutAuthor },
      message: {
        en: `${doc.undated} change(s) have no date and ${doc.withoutAuthor} have no author, so part of the timeline is lost.`,
        ja: `日時の無い変更が ${doc.undated} 件、作成者の無い変更が ${doc.withoutAuthor} 件あり、時系列の一部が失われています。`,
      },
    });
  }

  if (tpl) {
    // notFromTemplate
    if (tpl.rsidRoot && doc.rsidRoot && tpl.rsidRoot !== doc.rsidRoot) {
      evidence.push({
        id: "notFromTemplate",
        details: { templateRsidRoot: tpl.rsidRoot, rsidRoot: doc.rsidRoot },
        message: {
          en: "This file was not made from the template: its first editing session (rsidRoot) differs from the template's.",
          ja: "このファイルはテンプレートから作られていません。最初の編集セッションの識別子 (rsidRoot) がテンプレートと違います。",
        },
      });
    }

    // untrackedText
    const known = new Set([...tpl.listedRsids, ...tpl.usedRsids]);
    if (tpl.rsidRoot) known.add(tpl.rsidRoot);
    if (known.size && doc.usedRsids.size) {
      const sessions = new Set<string>();
      let chars = 0;
      for (const r of doc.runs) {
        if (r.tracked || !r.rsidR || known.has(r.rsidR) || r.chars === 0) continue;
        chars += r.chars;
        sessions.add(r.rsidR);
      }
      if (chars >= (opts.minUntrackedChars ?? 1)) {
        evidence.push({
          id: "untrackedText",
          details: { chars, sessions: sessions.size },
          message: {
            en:
              `${chars} character(s) of text, typed in ${sessions.size} editing session(s) that aren't in the template ` +
              "(or the previous submission), are not tracked changes. They were typed with Track Changes off, or their " +
              "changes were accepted.",
            ja:
              `テンプレート (または前回の提出物) に無い ${sessions.size} 回の編集セッションで入力された本文 ${chars} 文字が、` +
              "変更履歴になっていません。" +
              "記録をオフにして入力されたか、変更が承諾されています。",
          },
        });
      }
    }
  }

  return { checkedAgainstTemplate: !!tpl, evidence, note: TAMPER_NOTE };
}

/** 痕跡が見つかったか */
export function hasTamperEvidence(report: TamperReport | undefined): boolean {
  return !!report && report.evidence.length > 0;
}

/** 出力ファイル名に付ける印。痕跡が見つかった図は <名前>-tampered.svg にする */
export const TAMPERED_SUFFIX = "-tampered";

/** "<名前>.svg" を "<名前>-tampered.svg" にする */
export function tamperedFileName(file: string): string {
  return /\.svg$/i.test(file) ? file.replace(/(\.svg)$/i, `${TAMPERED_SUFFIX}$1`) : `${file}${TAMPERED_SUFFIX}`;
}

export function localize(text: LocalizedText, lang: Lang = getLang()): string {
  return typeof text === "string" ? text : text[lang] ?? text.en ?? Object.values(text)[0] ?? "";
}

const BANNER_TITLE: LocalizedText = {
  en: "Warning: traces of tampering with the tracked changes were found",
  ja: "警告: 変更履歴が改ざんされた痕跡が見つかりました",
};

function escXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** 全角は1、それ以外は0.55 として、幅 maxWidth に収まるように折り返す */
function wrap(text: string, fontSize: number, maxWidth: number): string[] {
  const lines: string[] = [];
  let line = "";
  let w = 0;
  for (const ch of text) {
    const cw = (ch.codePointAt(0)! >= 0x1100 ? 1 : 0.55) * fontSize;
    if (w + cw > maxWidth && line) {
      // 英文は単語の途中で切らない
      const sp = line.lastIndexOf(" ");
      if (ch !== " " && sp > 0 && /[\x21-\x7e]/.test(ch)) {
        lines.push(line.slice(0, sp));
        line = line.slice(sp + 1);
        w = [...line].reduce((s, c) => s + (c.codePointAt(0)! >= 0x1100 ? 1 : 0.55) * fontSize, 0);
      } else {
        lines.push(line);
        line = "";
        w = 0;
      }
      if (ch === " ") continue;
    }
    line += ch;
    w += cw;
  }
  if (line) lines.push(line);
  return lines;
}

const BANNER_FONT =
  "'Noto Sans CJK JP', 'Noto Sans JP', 'Yu Gothic', 'Hiragino Sans', Meiryo, 'MS PGothic', 'Helvetica Neue', Arial, sans-serif";

/**
 * 描いた SVG の上に、痕跡の警告を加える。元の図は <svg> のまま下にずらして入れ子にするため、
 * 図ごとのレイアウトには手を入れない。痕跡が無ければそのまま返す。
 */
export function addTamperWarning(svg: string, report: TamperReport | undefined, lang: Lang = getLang()): string {
  if (!hasTamperEvidence(report)) return svg;
  const root = svg.match(/<svg\b[^>]*>/);
  if (!root || root.index === undefined) return svg;
  const width = Number(root[0].match(/\bwidth="([\d.]+)"/)?.[1]);
  const height = Number(root[0].match(/\bheight="([\d.]+)"/)?.[1]);
  if (!Number.isFinite(width) || !Number.isFinite(height)) return svg;

  const pad = 12;
  const inner = width - pad * 4;
  const rows: { text: string; size: number; bold?: boolean; color: string }[] = [];
  for (const l of wrap(localize(BANNER_TITLE, lang), 15, inner)) rows.push({ text: l, size: 15, bold: true, color: "#9b1c1c" });
  for (const e of report!.evidence) {
    wrap(`・${localize(e.message, lang)}`, 12, inner).forEach((l, i) =>
      rows.push({ text: i === 0 ? l : `　${l}`, size: 12, color: "#1a1a1a" })
    );
  }
  for (const l of wrap(localize(report!.note, lang), 11, inner)) rows.push({ text: l, size: 11, color: "#555" });

  let y = pad * 2;
  const texts = rows.map((r) => {
    y += r.size + 5;
    return `<text x="${pad * 2}" y="${y}" font-size="${r.size}"${r.bold ? ' font-weight="bold"' : ""} fill="${r.color}">${escXml(r.text)}</text>`;
  });
  const bannerH = Math.ceil(y + pad * 2);
  const total = height + bannerH;

  const banner =
    `<g id="tamper-warning" font-family="${escXml(BANNER_FONT)}">` +
    `<rect x="${pad}" y="${pad}" width="${width - pad * 2}" height="${bannerH - pad * 2}" rx="6" fill="#fdecea" stroke="#d93025" stroke-width="2"/>` +
    texts.join("") +
    `</g>`;
  const ids = report!.evidence.map((e) => e.id).join(" ");
  const outerOpen =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${total}" viewBox="0 0 ${width} ${total}" data-tamper-evidence="${ids}">` +
    `\n  <rect width="100%" height="100%" fill="#ffffff"/>\n  ${banner}\n  `;
  // 元の <svg> 開始タグを、下にずらした位置に置き直す
  const innerOpen = root[0]
    .replace(/\swidth="[^"]*"/, "")
    .replace(/\sheight="[^"]*"/, "")
    .replace(/^<svg\b/, `<svg x="0" y="${bannerH}" width="${width}" height="${height}"`);
  const before = svg.slice(0, root.index);
  const after = svg.slice(root.index + root[0].length);
  return before + outerOpen + innerOpen + after.replace(/<\/svg>\s*$/, "</svg>\n</svg>\n");
}
