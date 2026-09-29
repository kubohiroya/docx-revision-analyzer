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
import * as fs from "fs";
import * as path from "path";
import { spawnSync } from "child_process";
import { formatTimestampForFilename } from "./filenames";

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
  return s.trackRevisions ? "" : "あわせて「変更履歴の記録」もオンにします。";
}

/**
 * 設定を有効化してよいかをユーザーに尋ねる文面 (ダイアログ用)。
 * 変更が不要な場合は undefined。
 */
export function buildEnableHistoryPrompt(s: DocxRevisionSettings): string | undefined {
  if (!needsHistoryFix(s)) return undefined;
  return (
    "この文書は編集者の名前と編集日時を保存するための設定が無効化されています。有効にしますか？\n\n" +
    "OK を選ぶと文書の設定を書き換えて上書き保存します (元のファイルはバックアップとして同じフォルダに残します)。" +
    trackRevisionsNote(s) +
    "Word でこの文書を開いている場合は、先に閉じてください。" +
    "すでに削除された過去の編集日時は復元できません。"
  );
}

/** 設定の問題を説明する文 (警告や y/N の質問の前置き)。変更が不要な場合は undefined */
export function describeHistorySettingsProblem(s: DocxRevisionSettings): string | undefined {
  if (!needsHistoryFix(s)) return undefined;
  return (
    "この文書は「保存時にファイルのプロパティから個人情報を削除する」設定が有効なため、" +
    "保存時に変更履歴の作成者と日時が削除されます。"
  );
}

/** --preserve-history で行う処理の説明 */
export function describePreserveHistoryAction(s: DocxRevisionSettings): string {
  return (
    "この設定を外し" +
    (s.trackRevisions ? "" : "、「変更履歴の記録」をオンにし") +
    "て上書き保存します (元のファイルはバックアップとして同じフォルダに残します)"
  );
}

/** CLI の警告として表示する文面。変更が不要な場合は undefined */
export function describeHistorySettingsWarning(s: DocxRevisionSettings): string | undefined {
  const problem = describeHistorySettingsProblem(s);
  if (!problem) return undefined;
  return `警告: ${problem}--preserve-history を指定すると、${describePreserveHistoryAction(s)}。`;
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
    throw new Error("settings.xml に WordprocessingML の名前空間が見つかりません。");
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
    throw new Error("settings.xml に w:settings 要素が見つかりません。");
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
    throw new Error("settings.xml の構造を解釈できませんでした。");
  }
  return out.slice(0, insertAt) + `<${w}:trackRevisions/>` + out.slice(insertAt);
}

// ---------------------------------------------------------------------------
// ファイルが他で開かれていないかの確認
// ---------------------------------------------------------------------------

/**
 * Word が文書を開いている間に同じフォルダへ作る所有者ファイル (~$xxxx.docx) を探す。
 * Word はファイル名が長い場合、先頭の1〜2文字を "~$" で置き換えた名前にする。
 */
function findWordLockFile(filePath: string): string | undefined {
  const dir = path.dirname(filePath);
  const name = path.basename(filePath).normalize("NFC");
  let entries: string[];
  try {
    entries = fs.readdirSync(dir);
  } catch {
    return undefined;
  }
  for (const entry of entries) {
    const e = entry.normalize("NFC");
    if (!e.startsWith("~$")) continue;
    const rest = e.slice(2);
    if (rest.length > 0 && rest.length >= name.length - 2 && name.endsWith(rest)) {
      return path.join(dir, entry);
    }
  }
  return undefined;
}

/** 他のプロセスがファイルを開いているか (OS ごとのベストエフォート) */
function isOpenByAnotherProcess(filePath: string): boolean {
  if (process.platform === "win32") {
    // Windows では Word が開いているファイルは共有違反で書き込みオープンできない
    try {
      fs.closeSync(fs.openSync(filePath, "r+"));
      return false;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      return code === "EBUSY" || code === "EPERM" || code === "EACCES";
    }
  }
  // macOS / Linux: lsof が使えれば、ファイルを開いているプロセスの有無を調べる
  const r = spawnSync("lsof", ["-t", "--", filePath], { encoding: "utf-8" });
  if (r.error) return false;
  return r.stdout.trim().length > 0;
}

/**
 * ファイルが Word 等で開かれていないことを確認する。
 * 開かれている可能性がある場合は、その理由を説明する文字列を返す。
 */
export function checkNotOpenElsewhere(filePath: string): string | undefined {
  const lock = findWordLockFile(filePath);
  if (lock) {
    return (
      `Word でこの文書が開かれています (所有者ファイル ${path.basename(lock)} があります)。` +
      "Word で文書を閉じてから再度実行してください。" +
      "Word を閉じてもこのファイルが残っている場合は、異常終了時の残骸なので削除してください。"
    );
  }
  if (isOpenByAnotherProcess(filePath)) {
    return "他のアプリケーションがこの文書を開いています。閉じてから再度実行してください。";
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// ファイルへの適用
// ---------------------------------------------------------------------------

export interface EnableHistoryResult {
  /** 設定を書き換えたか (もともと問題が無ければ false) */
  changed: boolean;
  /** 書き換え前のファイルのバックアップ先 */
  backupPath?: string;
  before: DocxRevisionSettings;
}

/**
 * 文書の設定を書き換え、変更履歴の作成者・日時が保存されるようにする。
 * 元のファイルは "<名前>.backup-<YYYYMMDD-HHMMSS>.docx" として同じフォルダに残す。
 * ファイルが他で開かれている場合はエラーを投げる。
 */
export async function enableHistoryPreservation(filePath: string): Promise<EnableHistoryResult> {
  const buf = await fs.promises.readFile(filePath);
  const zip = await JSZip.loadAsync(buf);
  const settingsFile = zip.file(SETTINGS_PART);
  const xml = settingsFile ? await settingsFile.async("string") : undefined;
  const before = parseHistorySettings(xml);
  if (xml === undefined || !needsHistoryFix(before)) return { changed: false, before };

  const openReason = checkNotOpenElsewhere(filePath);
  if (openReason) throw new Error(openReason);

  const patched = patchSettingsXml(xml);
  const after = parseHistorySettings(patched);
  if (needsHistoryFix(after) || !after.trackRevisions) {
    throw new Error("settings.xml の書き換えに失敗しました (書き換え後も設定が有効化されていません)。");
  }
  zip.file(SETTINGS_PART, patched);
  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });

  const dir = path.dirname(filePath);
  const base = path.basename(filePath).replace(/\.docx$/i, "");
  const backupPath = path.join(dir, `${base}.backup-${formatTimestampForFilename(new Date())}.docx`);
  await fs.promises.copyFile(filePath, backupPath, fs.constants.COPYFILE_EXCL);

  // 途中で失敗しても元のファイルが壊れないよう、一時ファイルに書いてから置き換える
  const tmpPath = path.join(dir, `.${base}.tmp-${process.pid}.docx`);
  try {
    await fs.promises.writeFile(tmpPath, out);
    await fs.promises.rename(tmpPath, filePath);
  } catch (err) {
    await fs.promises.rm(tmpPath, { force: true });
    throw err;
  }
  return { changed: true, backupPath, before };
}
