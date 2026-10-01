/**
 * historySettings.ts
 *
 * 変更履歴の時系列解析に必要な「文書ごとの設定」(word/settings.xml) の
 * 読み取り・診断・書き換えを行うモジュール。
 *
 * 対象とする設定:
 *  - w:trackRevisions            … 「変更履歴の記録」。無いと以後の編集が履歴に残らない。
 *  - w:removePersonalInformation … 「保存時にファイルのプロパティから個人情報を削除する」。
 *  - w:removeDateAndTime           有効だと Word は保存時に変更履歴の作成者と日時 (w:date) を消す。
 *
 * macOS 版 Word には後者2つを切り替える UI が無いため、本モジュールで
 * settings.xml を直接書き換えられるようにしている。
 */

import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { t } from "./i18n";
import type { DocxInput } from "./input";

export interface DocxRevisionSettings {
  /** word/settings.xml が存在するか (無い場合、以下はすべて false で書き換えもできない) */
  settingsPartFound: boolean;
  /** 「変更履歴の記録」が有効な状態で保存されたか (w:trackRevisions) */
  trackRevisions: boolean;
  /** 「保存時にファイルのプロパティから個人情報を削除する」(w:removePersonalInformation) */
  removePersonalInformation: boolean;
  /** 変更履歴・コメントの日時を削除する設定 (w:removeDateAndTime) */
  removeDateAndTime: boolean;
}

const SETTINGS_PART = "word/settings.xml";

type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
});

/** w:settings 直下の要素を探し、ST_OnOff 型として真偽を返す (要素が無ければ false) */
function readOnOff(settingsChildren: XmlNode[], name: string): boolean {
  for (const node of settingsChildren) {
    if (!(name in node)) continue;
    const attrs = node[":@"] as Record<string, unknown> | undefined;
    const val = attrs?.["@_val"];
    if (val === undefined) return true;
    return !["false", "0", "off"].includes(String(val).toLowerCase());
  }
  return false;
}

/** settings.xml の文字列から設定を読み取る */
export function parseHistorySettings(settingsXml: string | undefined): DocxRevisionSettings {
  if (!settingsXml) {
    return {
      settingsPartFound: false,
      trackRevisions: false,
      removePersonalInformation: false,
      removeDateAndTime: false,
    };
  }
  const root = parser.parse(settingsXml) as XmlNode[];
  const settingsNode = root.find((n) => "settings" in n);
  const children = (settingsNode?.["settings"] as XmlNode[] | undefined) ?? [];
  return {
    settingsPartFound: true,
    trackRevisions: readOnOff(children, "trackRevisions"),
    removePersonalInformation: readOnOff(children, "removePersonalInformation"),
    removeDateAndTime: readOnOff(children, "removeDateAndTime"),
  };
}

export async function readHistorySettingsFromZip(zip: JSZip): Promise<DocxRevisionSettings> {
  const file = zip.file(SETTINGS_PART);
  return parseHistorySettings(file ? await file.async("string") : undefined);
}

/** 作成者・日時が保存時に削除される設定になっているか */
export function removesAuthorAndDate(s: DocxRevisionSettings): boolean {
  return s.removePersonalInformation || s.removeDateAndTime;
}

/**
 * 設定の書き換えを提案すべきか。
 * 作成者・日時が削除される設定のときだけ対象にする (「変更履歴の記録」がオフなだけの文書は、
 * 受け取った文書を解析するだけの利用者にとって煩わしいため尋ねない)。
 * settings.xml が無い文書は書き換えられないため対象外。
 */
export function needsHistoryFix(s: DocxRevisionSettings): boolean {
  return s.settingsPartFound && removesAuthorAndDate(s);
}

/** 書き換え時にあわせて「変更履歴の記録」もオンにすることを伝える文 (既にオンなら空文字列) */
function trackRevisionsNote(s: DocxRevisionSettings): string {
  return s.trackRevisions ? "" : t("trackRevisionsNote");
}

/**
 * 設定を有効化してよいかをユーザーに尋ねる文面 (ダイアログ用)。
 * 変更が不要な場合は undefined。
 */
export function buildEnableHistoryPrompt(s: DocxRevisionSettings): string | undefined {
  if (!needsHistoryFix(s)) return undefined;
  return t("enableHistoryPrompt", trackRevisionsNote(s));
}

/** 設定の問題を説明する文 (警告や y/N の質問の前置き)。変更が不要な場合は undefined */
export function describeHistorySettingsProblem(s: DocxRevisionSettings): string | undefined {
  if (!needsHistoryFix(s)) return undefined;
  return t("settingsProblem");
}

/** --preserve-history で行う処理の説明 */
export function describePreserveHistoryAction(s: DocxRevisionSettings): string {
  return t("preserveAction", s.trackRevisions);
}

/** CLI の警告として表示する文面。変更が不要な場合は undefined */
export function describeHistorySettingsWarning(s: DocxRevisionSettings): string | undefined {
  const problem = describeHistorySettingsProblem(s);
  if (!problem) return undefined;
  return t("settingsWarning", problem, describePreserveHistoryAction(s));
}

// ---------------------------------------------------------------------------
// settings.xml の書き換え
// ---------------------------------------------------------------------------

const WML_NAMESPACES = [
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
];

/**
 * CT_Settings のスキーマ上、w:trackRevisions より後ろに置かれる要素。
 * Word は要素順がスキーマと異なると「ファイルが破損している」と判定することがあるため、
 * これらのうち最初に現れる要素の直前に w:trackRevisions を挿入する。
 */
const ELEMENTS_AFTER_TRACK_REVISIONS = new Set([
  "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride",
  "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation",
  "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope",
  "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders",
  "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets",
  "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
  "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
  "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin",
  "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning",
  "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars",
  "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture",
  "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent",
  "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly",
  "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
  "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars",
  "rsids", "attachedSchema", "themeFontLang", "clrSchemeMapping",
  "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions",
  "readModeInkLockDown", "smartTagType", "shapeDefaults", "doNotEmbedSmartTags",
  "decimalSymbol", "listSeparator",
]);

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** 空要素 (<p:name .../>) と中身付き要素 (<p:name ...>...</p:name>) の両方にマッチする正規表現 */
function elementRegExp(prefix: string, name: string): RegExp {
  const q = escapeRegExp(`${prefix}:${name}`);
  return new RegExp(`<${q}(?=[\\s/>])[^>]*?/>|<${q}(?=[\\s>])[^>]*>[\\s\\S]*?</${q}\\s*>`, "g");
}

/**
 * settings.xml の文字列を書き換え、個人情報削除の設定を外して変更履歴の記録を有効にする。
 * 書式を保つため XML を再シリアライズせず、該当要素だけを文字列として操作する。
 */
export function patchSettingsXml(xml: string): string {
  const nsMatch = WML_NAMESPACES.map((ns) =>
    xml.match(new RegExp(`xmlns:([A-Za-z_][\\w.-]*)="${escapeRegExp(ns)}"`))
  ).find((m) => m);
  if (!nsMatch) {
    throw new Error(t("errNoWmlNamespace"));
  }
  const w = nsMatch[1];

  let out = xml
    .replace(elementRegExp(w, "removePersonalInformation"), "")
    .replace(elementRegExp(w, "removeDateAndTime"), "");

  const trackRe = elementRegExp(w, "trackRevisions");
  if (trackRe.test(out)) {
    // w:val="false" 等で明示的にオフになっている場合も含め、オンの形に置き換える
    out = out.replace(elementRegExp(w, "trackRevisions"), `<${w}:trackRevisions/>`);
    return out;
  }

  // w:settings の開始タグの後ろから、trackRevisions より後ろに置くべき最初の要素を探す
  const rootOpen = out.match(new RegExp(`<${escapeRegExp(w)}:settings(?=[\\s>])[^>]*>`));
  if (!rootOpen || rootOpen.index === undefined) {
    throw new Error(t("errNoSettingsElement"));
  }
  const bodyStart = rootOpen.index + rootOpen[0].length;
  const tagRe = /<(\/?)([A-Za-z_][\w.-]*):([A-Za-z_][\w.-]*)/g;
  tagRe.lastIndex = bodyStart;
  let insertAt = -1;
  for (let m = tagRe.exec(out); m; m = tagRe.exec(out)) {
    const [, closing, prefix, local] = m;
    if (closing) {
      if (prefix === w && local === "settings") {
        insertAt = m.index;
        break;
      }
      continue;
    }
    // w 以外の名前空間の要素 (m:mathPr, w14:*, w15:* など) はすべて trackRevisions より後ろ
    if (prefix !== w || ELEMENTS_AFTER_TRACK_REVISIONS.has(local)) {
      insertAt = m.index;
      break;
    }
  }
  if (insertAt < 0) {
    throw new Error(t("errSettingsStructure"));
  }
  return out.slice(0, insertAt) + `<${w}:trackRevisions/>` + out.slice(insertAt);
}

// ---------------------------------------------------------------------------
// バイト列の書き換え
// ---------------------------------------------------------------------------

export interface PreserveHistoryResult {
  /** 設定を書き換えたか (もともと問題が無ければ false) */
  changed: boolean;
  before: DocxRevisionSettings;
  /** 書き換えた docx のバイト列 (changed のときだけ) */
  output?: Uint8Array;
}

/** docx のバイト列から変更履歴に関する設定を読み取る */
export async function readHistorySettings(input: DocxInput): Promise<DocxRevisionSettings> {
  return readHistorySettingsFromZip(await JSZip.loadAsync(input));
}

/**
 * docx のバイト列の設定を書き換え、変更履歴の作成者・日時が保存されるようにした docx を返す。
 * ファイルシステムは使わない (ファイルへの保存は node/historyFile.ts の enableHistoryPreservation)。
 */
export async function preserveHistoryInDocx(input: DocxInput): Promise<PreserveHistoryResult> {
  const zip = await JSZip.loadAsync(input);
  const settingsFile = zip.file(SETTINGS_PART);
  const xml = settingsFile ? await settingsFile.async("string") : undefined;
  const before = parseHistorySettings(xml);
  if (xml === undefined || !needsHistoryFix(before)) return { changed: false, before };

  const patched = patchSettingsXml(xml);
  const after = parseHistorySettings(patched);
  if (needsHistoryFix(after) || !after.trackRevisions) {
    throw new Error(t("errPatchFailed"));
  }
  zip.file(SETTINGS_PART, patched);
  const output = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { changed: true, before, output };
}
