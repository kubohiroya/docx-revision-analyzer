/**
 * integrity.ts
 *
 * 変更履歴の「整合性の簡易チェック」。判定や警告には使わず、情報として出す。
 *
 * ローカルに保存された docx の変更履歴は、記録をオフにする・すべて承諾する・XML を直接編集して
 * w:date / w:author を書き換える、といった操作で消したり作り替えたりできる。ここでは、Word が通常
 * 書き出す形と違う点を数え上げるだけで、それが改変によるものかどうかは判断しない
 * (Word 以外のアプリ・古い Word・変換ツールでも同じ特徴は現れる)。
 *
 * 項目:
 *  - settings: 変更履歴の記録のオン/オフ、保存時に作成者・日時を削除する設定
 *  - undated: 日時・作成者の無い変更の数
 *  - duplicateIds: 変更 (w:ins / w:del / w:moveFrom / w:moveTo / 書式の変更) の w:id の重複
 *  - idOrder: w:id が文書の順に増えていない箇所の数 (Word は保存時に文書の順に番号を振り直す)
 *  - futureDates: 現在時刻や、文書の最終更新日時 (docProps/core.xml) より後の日時
 *  - beforeCreated: 文書の作成日時より前の日時
 *  - deletedBeforeInserted: 挿入より前の日時に削除された文章
 *  - rsids: 編集セッションの識別子 (rsid) が、settings.xml の一覧に無いのに本文で使われている数
 *  - application: 最後に保存したアプリ (docProps/app.xml)
 */

import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type { DocxInput } from "./input";
import type { LocalizedText } from "./categories";
import { parseDocxLayout } from "./docxLayout";
import { parseHistorySettings } from "./historySettings";

export interface IntegrityItem {
  id:
    | "settings"
    | "undated"
    | "duplicateIds"
    | "idOrder"
    | "futureDates"
    | "beforeCreated"
    | "deletedBeforeInserted"
    | "rsids"
    | "application";
  /**
   * Word が通常書き出す形と違う点が見つかったか。調べられなかった (必要な情報が文書に無い) 場合は null。
   * 判定ではなく、振り返りの手がかりとして示す
   */
  observed: boolean | null;
  /** 数値などの詳細 */
  details: Record<string, string | number | boolean | null>;
  message: LocalizedText;
}

export interface IntegrityReport {
  items: IntegrityItem[];
  /** 説明 (判定ではないこと) */
  note: LocalizedText;
}

export const INTEGRITY_NOTE: LocalizedText = {
  en:
    "Informational only. A .docx saved on a computer can have its tracked changes turned off, accepted, or rewritten " +
    "(including dates and authors), and other apps and converters also write files differently from Word. These " +
    "observations are not a judgment about how the document was written.",
  ja:
    "情報として示すだけです。コンピュータに保存された .docx の変更履歴は、記録をオフにする・すべて承諾する・日時や作成者を" +
    "書き換えるといった操作で消したり作り替えたりでき、Word 以外のアプリや変換ツールも Word と違う形で書き出します。" +
    "ここに示す点は、文書がどう書かれたかについての判定ではありません。",
};

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

const REVISION_TAGS = new Set([
  "ins",
  "del",
  "moveFrom",
  "moveTo",
  "rPrChange",
  "pPrChange",
  "sectPrChange",
  "tblPrChange",
  "trPrChange",
  "tcPrChange",
  "tblGridChange",
  "numberingChange",
  "cellIns",
  "cellDel",
]);

/** 本文 (段落・表) の変更。書式の変更は日時の比較には使う */
interface RevisionAttr {
  tag: string;
  id?: string;
  date?: string;
  author?: string;
}

function tagOf(n: XmlNode): string | undefined {
  return Object.keys(n).find((k) => k !== ":@" && k !== "#text");
}

function attrs(n: XmlNode): Record<string, string> {
  const a = (n[":@"] as Record<string, unknown> | undefined) ?? {};
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(a)) out[k.replace(/^@_/, "")] = String(v);
  return out;
}

/** 文書順に、変更の要素と rsid の属性を集める */
function walk(nodes: XmlNode[], revs: RevisionAttr[], rsids: Set<string>): void {
  for (const n of nodes) {
    const tag = tagOf(n);
    if (!tag) continue;
    const a = attrs(n);
    for (const [k, v] of Object.entries(a)) if (/^rsid/.test(k) && /^[0-9A-Fa-f]{8}$/.test(v)) rsids.add(v.toUpperCase());
    // w:ins / w:del は、段落記号の変更 (rPr の中) にも現れる。id のある要素だけを変更として数える
    if (REVISION_TAGS.has(tag) && a.id !== undefined) revs.push({ tag, id: a.id, date: a.date, author: a.author });
    const kids = n[tag];
    if (Array.isArray(kids)) walk(kids as XmlNode[], revs, rsids);
  }
}

function textOfPart(xml: string | undefined, tag: string): string | undefined {
  if (!xml) return undefined;
  const m = xml.match(new RegExp(`<(?:[\\w]+:)?${tag}\\b[^>]*>([^<]*)</(?:[\\w]+:)?${tag}>`));
  return m ? m[1].trim() : undefined;
}

const DAY_MS = 24 * 3600 * 1000;
/** 文書の最終更新日時と変更の日時を比べるときの余裕 (保存にかかる時間・時計のずれ) */
const MODIFIED_TOLERANCE_MS = 10 * 60 * 1000;

function fmt(d: Date | undefined): string | null {
  return d ? d.toISOString() : null;
}

/** docx のバイト列から、整合性の簡易チェックを行う */
export async function checkIntegrity(input: DocxInput, opts: { now?: Date } = {}): Promise<IntegrityReport> {
  const now = opts.now ?? new Date();
  const zip = await JSZip.loadAsync(input);
  const read = async (p: string) => (zip.file(p) ? await zip.file(p)!.async("string") : undefined);
  const [docXml, settingsXml, coreXml, appXml] = await Promise.all([
    read("word/document.xml"),
    read("word/settings.xml"),
    read("docProps/core.xml"),
    read("docProps/app.xml"),
  ]);
  const revs: RevisionAttr[] = [];
  const usedRsids = new Set<string>();
  if (docXml) walk(parser.parse(docXml) as XmlNode[], revs, usedRsids);
  const items: IntegrityItem[] = [];

  // settings
  const settings = parseHistorySettings(settingsXml);
  items.push({
    id: "settings",
    observed: settings.settingsPartFound ? !settings.trackRevisions || settings.removePersonalInformation || settings.removeDateAndTime : null,
    details: { ...settings },
    message: !settings.settingsPartFound
      ? { en: "The document has no settings part (word/settings.xml).", ja: "文書に設定 (word/settings.xml) がありません。" }
      : {
          en:
            `Track Changes was ${settings.trackRevisions ? "on" : "off"} when the document was saved` +
            (settings.removePersonalInformation || settings.removeDateAndTime
              ? "; Word is set to remove authors and dates of changes on save."
              : "."),
          ja:
            `保存時の「変更履歴の記録」は${settings.trackRevisions ? "オン" : "オフ"}でした` +
            (settings.removePersonalInformation || settings.removeDateAndTime
              ? "。保存時に変更の作成者・日時を削除する設定になっています。"
              : "。"),
        },
  });

  // undated
  const textRevs = revs.filter((r) => ["ins", "del", "moveFrom", "moveTo"].includes(r.tag));
  const undated = textRevs.filter((r) => !r.date).length;
  const noAuthor = textRevs.filter((r) => !r.author).length;
  const authors = new Set(textRevs.map((r) => r.author).filter(Boolean));
  items.push({
    id: "undated",
    observed: textRevs.length ? undated > 0 || noAuthor > 0 : null,
    details: { revisions: textRevs.length, undated, withoutAuthor: noAuthor, authors: authors.size },
    message: {
      en: `${textRevs.length} tracked insertions/deletions by ${authors.size} author(s); ${undated} without a date, ${noAuthor} without an author.`,
      ja: `変更 (挿入・削除) は ${textRevs.length} 件、作成者は ${authors.size} 人。日時の無いものが ${undated} 件、作成者の無いものが ${noAuthor} 件。`,
    },
  });

  // duplicateIds
  const idCount = new Map<string, number>();
  for (const r of revs) idCount.set(r.id!, (idCount.get(r.id!) ?? 0) + 1);
  const dupIds = [...idCount.values()].filter((n) => n > 1).length;
  items.push({
    id: "duplicateIds",
    observed: revs.length ? dupIds > 0 : null,
    details: { revisionElements: revs.length, duplicatedIds: dupIds },
    message: dupIds
      ? {
          en: `${dupIds} change id(s) (w:id) are used more than once. Word normally gives each change its own id.`,
          ja: `${dupIds} 個の変更の id (w:id) が重複しています。Word は通常、変更ごとに別の id を付けます。`,
        }
      : { en: "Each change has its own id (w:id).", ja: "変更の id (w:id) に重複はありません。" },
  });

  // idOrder
  let decreases = 0;
  const numeric = revs.map((r) => Number(r.id)).filter((n) => Number.isFinite(n));
  for (let i = 1; i < numeric.length; i++) if (numeric[i] < numeric[i - 1]) decreases++;
  items.push({
    id: "idOrder",
    observed: numeric.length > 1 ? decreases > 0 : null,
    details: { decreases },
    message: decreases
      ? {
          en: `Change ids decrease ${decreases} time(s) in document order. Word renumbers changes in document order when it saves.`,
          ja: `文書の順に見て、変更の id が小さくなる箇所が ${decreases} か所あります。Word は保存時に文書の順に番号を振り直します。`,
        }
      : { en: "Change ids increase in document order.", ja: "変更の id は文書の順に増えています。" },
  });

  // futureDates / beforeCreated
  const created = textOfPart(coreXml, "created");
  const modified = textOfPart(coreXml, "modified");
  const createdAt = created ? new Date(created) : undefined;
  const modifiedAt = modified ? new Date(modified) : undefined;
  const dates = revs.map((r) => (r.date ? new Date(r.date) : undefined)).filter((d): d is Date => !!d && !Number.isNaN(d.getTime()));
  const afterNow = dates.filter((d) => d.getTime() > now.getTime() + DAY_MS).length;
  const afterModified = modifiedAt ? dates.filter((d) => d.getTime() > modifiedAt.getTime() + MODIFIED_TOLERANCE_MS).length : 0;
  items.push({
    id: "futureDates",
    observed: dates.length ? afterNow > 0 || afterModified > 0 : null,
    details: { afterNow, afterLastModified: afterModified, lastModified: fmt(modifiedAt) },
    message: {
      en:
        (afterNow ? `${afterNow} change date(s) are in the future` : "No change date is in the future") +
        (modifiedAt
          ? `; ${afterModified} are later than the document's last-modified time (${modifiedAt.toISOString()}).`
          : "; the document has no last-modified time."),
      ja:
        (afterNow ? `未来の日時の変更が ${afterNow} 件あります` : "未来の日時の変更はありません") +
        (modifiedAt
          ? `。文書の最終更新日時 (${modifiedAt.toISOString()}) より後の変更は ${afterModified} 件です。`
          : "。文書に最終更新日時がありません。"),
    },
  });
  const beforeCreated = createdAt ? dates.filter((d) => d.getTime() < createdAt.getTime() - DAY_MS).length : 0;
  items.push({
    id: "beforeCreated",
    observed: createdAt && dates.length ? beforeCreated > 0 : null,
    details: { beforeCreated, created: fmt(createdAt) },
    message: createdAt
      ? {
          en: `${beforeCreated} change date(s) are more than a day before the document was created (${createdAt.toISOString()}). Copying text from another tracked document can also cause this.`,
          ja: `文書の作成日時 (${createdAt.toISOString()}) より1日以上前の日時の変更が ${beforeCreated} 件あります。変更履歴の付いた別の文書から文章を写した場合にも起こります。`,
        }
      : { en: "The document has no creation time.", ja: "文書に作成日時がありません。" },
  });

  // deletedBeforeInserted
  let reversed = 0;
  let both = 0;
  if (docXml) {
    const model = await parseDocxLayout(input);
    for (const p of model.paras) {
      for (const s of p.segs) {
        if (!s.ins?.date || !s.del?.date) continue;
        both++;
        if (s.del.date.getTime() < s.ins.date.getTime()) reversed++;
      }
    }
  }
  items.push({
    id: "deletedBeforeInserted",
    observed: both ? reversed > 0 : null,
    details: { insertedThenDeleted: both, deletedBeforeInserted: reversed },
    message: both
      ? {
          en: `${reversed} of ${both} passages that were inserted and later deleted have a deletion time earlier than the insertion time.`,
          ja: `挿入した後に削除された文章 ${both} 件のうち、削除の日時が挿入の日時より前のものが ${reversed} 件あります。`,
        }
      : { en: "No passage was inserted and later deleted.", ja: "挿入した後に削除された文章はありません。" },
  });

  // rsids
  const listed = new Set((settingsXml?.match(/<w:rsid\b[^>]*w:val="([0-9A-Fa-f]{8})"/g) ?? []).map((m) => m.slice(-9, -1).toUpperCase()));
  const root = settingsXml?.match(/<w:rsidRoot\b[^>]*w:val="([0-9A-Fa-f]{8})"/)?.[1]?.toUpperCase();
  if (root) listed.add(root);
  const unlisted = [...usedRsids].filter((r) => !listed.has(r)).length;
  items.push({
    id: "rsids",
    observed: listed.size && usedRsids.size ? unlisted > 0 : null,
    details: { listedInSettings: listed.size, usedInDocument: usedRsids.size, usedButNotListed: unlisted },
    message:
      listed.size && usedRsids.size
        ? {
            en: `Word recorded ${listed.size} editing-session id(s) (rsid); ${unlisted} rsid(s) used in the text are not in that list.`,
            ja: `Word が記録した編集セッションの識別子 (rsid) は ${listed.size} 個。本文で使われている rsid のうち ${unlisted} 個がその一覧にありません。`,
          }
        : {
            en: "The document has no editing-session ids (rsid). Apps other than Word often don't write them.",
            ja: "文書に編集セッションの識別子 (rsid) がありません。Word 以外のアプリは書き出さないことが多いです。",
          },
  });

  // application
  const application = textOfPart(appXml, "Application");
  const appVersion = textOfPart(appXml, "AppVersion");
  const totalTime = textOfPart(appXml, "TotalTime");
  items.push({
    id: "application",
    observed: null,
    details: { application: application ?? null, appVersion: appVersion ?? null, totalEditingMinutes: totalTime ? Number(totalTime) : null },
    message: application
      ? {
          en: `Last saved by ${application}${appVersion ? ` ${appVersion}` : ""}${totalTime ? `; total editing time recorded: ${totalTime} min` : ""}.`,
          ja: `最後に保存したアプリ: ${application}${appVersion ? ` ${appVersion}` : ""}${totalTime ? `。記録された編集時間の合計: ${totalTime} 分` : ""}。`,
        }
      : { en: "The document doesn't record which app saved it.", ja: "どのアプリで保存したかが記録されていません。" },
  });

  return { items, note: INTEGRITY_NOTE };
}
