/**
 * i18n.ts
 *
 * 表示言語 (英語 / 日本語) の判定とメッセージカタログ。既定は英語。
 *
 * 言語は次の順で決める (最初に見つかったもの):
 *   1. --lang <en|ja> (CLI の引数。initLangFromArgv で読む)
 *   2. 設定ファイル (<ツール名>.yml) の lang
 *   3. 環境変数 DOCX_REVISION_LANG
 *   4. 環境変数 LC_ALL / LC_MESSAGES / LANG / LANGUAGE
 *   5. macOS ではシステムの言語設定 (defaults read -g AppleLocale)、
 *      それ以外は Intl の既定ロケール
 * いずれも日本語 (ja...) でなければ英語とする。
 */

import { spawnSync } from "child_process";

export type Lang = "en" | "ja";

export const SUPPORTED_LANGS: readonly Lang[] = ["en", "ja"];

/** ロケール文字列 (ja_JP.UTF-8, ja-JP, en_US など) を Lang に変換する。判定できなければ undefined */
export function langFromLocale(locale: string | undefined): Lang | undefined {
  if (!locale) return undefined;
  const v = locale.trim().toLowerCase();
  if (v === "" || v === "c" || v === "posix" || v.startsWith("c.")) return undefined;
  return v.startsWith("ja") ? "ja" : "en";
}

function systemLocale(): string | undefined {
  if (process.platform === "darwin") {
    try {
      const r = spawnSync("defaults", ["read", "-g", "AppleLocale"], { encoding: "utf-8", timeout: 2000 });
      if (r.status === 0 && r.stdout.trim()) return r.stdout.trim();
    } catch {
      // 取得できなければ Intl に任せる
    }
  }
  try {
    return Intl.DateTimeFormat().resolvedOptions().locale;
  } catch {
    return undefined;
  }
}

/** 環境 (環境変数・OS の設定) から表示言語を判定する */
export function detectLang(env: NodeJS.ProcessEnv = process.env): Lang {
  const candidates = [
    env.DOCX_REVISION_LANG,
    env.LC_ALL,
    env.LC_MESSAGES,
    env.LANG,
    env.LANGUAGE?.split(":")[0],
  ];
  for (const c of candidates) {
    const lang = langFromLocale(c);
    if (lang) return lang;
  }
  return langFromLocale(systemLocale()) ?? "en";
}

let current: Lang | undefined;

export function setLang(lang: Lang): void {
  current = lang;
}

export function getLang(): Lang {
  if (!current) current = detectLang();
  return current;
}

/**
 * CLI の引数から --lang / --lang=xx を読み、なければ環境から判定して言語を確定する。
 * コマンドのヘルプ文は定義時に作られるため、commander に渡す前に呼ぶ。
 */
export function initLangFromArgv(argv: string[] = process.argv, configLang?: unknown): Lang {
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const value = a === "--lang" ? argv[i + 1] : a.startsWith("--lang=") ? a.slice("--lang=".length) : undefined;
    if (value !== undefined) {
      const lang = SUPPORTED_LANGS.find((l) => l === value.toLowerCase()) ?? langFromLocale(value) ?? "en";
      setLang(lang);
      return lang;
    }
  }
  const fromConfig = typeof configLang === "string" ? SUPPORTED_LANGS.find((l) => l === configLang.toLowerCase()) : undefined;
  const lang = fromConfig ?? detectLang();
  setLang(lang);
  return lang;
}

// ---------------------------------------------------------------------------
// メッセージカタログ
// ---------------------------------------------------------------------------

const en = {
  // 共通
  errorFor: (file: string, msg: string) => `Error (${file}): ${msg}`,
  doneFor: (file: string, msg: string) => `Done (${file}): ${msg}`,
  warning: (msg: string) => `Warning: ${msg}`,
  fileNotFound: (p: string) => `File not found: ${p}`,
  errMultiOutput: "Error: -o/--output can't be used with multiple input files.",
  errPositive: (name: string) => `${name} must be a positive number.`,
  optLang: "Display language (en / ja). Defaults to the OS locale",
  configLoaded: (p: string) => `Using settings from ${p}`,
  configInvalid: (p: string, msg: string) => `Couldn't read the settings file ${p}: ${msg}`,
  configNotMapping: "it must be a list of key: value pairs",
  configUnknownKey: (key: string, p: string) => `Ignoring unknown key "${key}" in ${p}`,
  configOutputIgnored: (p: string) => `Ignoring "output" in ${p} because several files were given`,
  configBadValue: (key: string, value: string, p: string) => `Ignoring invalid value "${value}" for "${key}" in ${p}`,
  optDrop:
    "Desktop drag-and-drop mode. Without -o, the output file name gets the input file's last-modified time " +
    "(for launches without a terminal, such as the macOS Finder droplet or dropping files on the .exe in Windows " +
    "Explorer; on Windows the result is also shown in a message box)",
  optPreserveHistory:
    "If the document is set to remove personal information (tracked-change authors and dates) on save, turn that " +
    "off, turn Track Changes on, and save it in place (the original is kept as <name>.backup-<time>.docx). " +
    "Fails without changing anything if the document is open in Word or another app",
  optPreserveHistoryAlias: "Alias of --preserve-history",
  optCheckHistorySettings:
    "Don't produce output; check each file's settings and print the result to stdout. The first line is ok or " +
    "needs-fix, followed by the confirmation text for needs-fix (used by the droplet)",
  optBulkChars:
    "Treat insertions by the same author at the same time totalling at least this many characters as a bulk " +
    "insertion (orange)",
  defaultTitle: (prefix: string, name: string, time: string) => `${prefix}: ${name} (last modified ${time})`,
  titleTime: (y: number, mo: string, d: string, h: string, mi: string) => `${y}-${mo}-${d} ${h}:${mi}`,

  // --preserve-history
  preservePrompt: (file: string, problem: string, action: string) =>
    `${file}: ${problem}\nRun the --preserve-history fix: ${action}. If the document is open in Word, close it first.\nProceed? [y/N] `,
  preserveDone: (backup: string) =>
    "Rewrote the settings to keep tracked-change authors and dates and to track changes, and saved the document " +
    `(original: ${backup}). Edits from now on will be timestamped.`,
  trackRevisionsNote: "Track Changes will be turned on as well. ",
  enableHistoryPrompt: (trackNote: string) =>
    "This document is set to discard editor names and edit times. Enable saving them?\n\n" +
    "Choosing OK rewrites the document's settings and saves it in place (the original is kept as a backup in the " +
    `same folder). ${trackNote}If the document is open in Word, close it first. ` +
    "Edit times that were already removed can't be recovered.",
  settingsProblem:
    'This document has Word\'s "Remove personal information from file properties on save" option on, so ' +
    "tracked-change authors and dates are removed whenever it's saved. ",
  preserveAction: (trackRevisions: boolean) =>
    `turn that option off${trackRevisions ? "" : " and turn Track Changes on"}, then save the document in place ` +
    "(the original is kept as a backup in the same folder)",
  settingsWarning: (problem: string, action: string) => `Warning: ${problem}With --preserve-history, the tool will ${action}.`,
  errNoWmlNamespace: "settings.xml has no WordprocessingML namespace.",
  errNoSettingsElement: "settings.xml has no w:settings element.",
  errSettingsStructure: "Couldn't interpret the structure of settings.xml.",
  errPatchFailed: "Failed to rewrite settings.xml (the settings are still not enabled after the rewrite).",
  openInWord: (lock: string) =>
    `The document is open in Word (owner file ${lock} exists). Close it in Word and try again. ` +
    "If that file is still there after closing Word, it was left behind by a crash; delete it.",
  openElsewhere: "Another application has this document open. Close it and try again.",

  // 変更履歴が取れない理由
  undatedCause:
    'Word\'s "Remove personal information from file properties on save" option is on, so the dates were removed. ' +
    "Removed dates can't be recovered.",
  undatedRevisions: (n: number, cause: string) =>
    `Found ${n} tracked changes (w:ins / w:del), but none has a timestamp (w:date), so there's no timeline to analyze. ${cause}`,
  noRevisionsTrackingOff:
    "No tracked changes (w:ins / w:del) were found. This file was saved with Track Changes off. " +
    "Tracked changes also disappear once they're all accepted.",
  noRevisionsTrackingOn:
    "No tracked changes (w:ins / w:del) were found. Track Changes is on, but there are no changes yet, " +
    "or they have all been accepted or rejected.",
  errNoDocumentXml: "word/document.xml not found.",
  errNoBody: "word/document.xml has no body (w:body).",
  noInsertions: "No insertion events were found, so no score can be computed (returning 0).",
  fewEvents: "Fewer than 5 insertion events, so the result is statistically weak. Treat it as a rough indication.",

  // docx-revision-chart
  chartTitlePrefix: "Revision history",
  chartDescription:
    "Generate an SVG chart of characters added/deleted over time and the total character count from a Word file " +
    "(.docx) edited with Track Changes. Multiple files are processed one by one; a failure doesn't stop the rest.",
  chartArgFiles: ".docx files to analyze (several allowed, e.g. files dropped onto the .exe in Windows Explorer)",
  chartOptOutput: "Output SVG path (default: same name as the input with .svg). Not allowed with multiple files",
  chartOptBucket: "Time bucket size: auto|second|minute|hour|day, or a number of seconds",
  chartOptGap:
    "Threshold (hours) separating periods of continuous editing from idle periods. When set, draws a separate " +
    "chart for each period, side by side, with a gap showing the length of each idle period",
  chartOptWidth: "SVG width (px) (by default computed from the content when -p is used)",
  chartOptHeight: "SVG height (px)",
  chartOptTitle: "Chart title",
  errGapThreshold: "--gap-threshold (-p) must be a number of hours of 0 or more.",
  chartDoneSessions: (events: number, sessions: number, out: string) =>
    `Split ${events} events into ${sessions} periods and wrote ${out}.`,
  chartDoneBuckets: (buckets: number, out: string) => `Aggregated into ${buckets} time buckets and wrote ${out}.`,
  kindFine: "Fine-grained edits",
  kindBulk: "Bulk insertion",
  kindMoved: "Moved / reordered",
  kindDeleted: "Deleted",
  totalChars: "Total characters",
  barTooltip: (label: string, n: number, when: string) => `${label} ${n} chars (${when})`,
  totalTooltip: (n: number, when: string) => `Total ${n} chars (${when})`,
  axisAddedDeleted: "Chars (added/deleted)",
  chartDefaultTitle: "Revision history (added/deleted characters, total characters)",
  chartDefaultTitleSessions: "Revision history by period (added/deleted characters, total characters)",
  chartSplitNote: (hours: number) => `(split at idle periods of ${hours}+ hours)`,
  errNoBuckets: "No tracked changes to draw (buckets is empty).",
  errNoSessions: "No tracked changes to draw (sessions is empty).",

  // docx-revision-flow
  flowTitlePrefix: "Edit flow",
  flowDescription:
    "Split a Word file's (.docx) Track Changes into sessions of continuous editing and draw, left to right, the " +
    "document at the start of the first session and at the end of each session as schematic page thumbnails. " +
    "Paragraphs edited in fine steps are green, bulk-inserted or replaced ones orange, and bands between the " +
    "columns show how each paragraph/figure was deleted, replaced, grown, or moved during the session (one SVG).",
  flowArgFiles: ".docx files to analyze (several allowed)",
  flowOptOutput: "Output SVG path (default: <input name>-flow.svg). Not allowed with multiple files",
  flowOptGap: "Start a new session after an idle period longer than this many hours",
  flowOptFrom: 'Start of the period to analyze (e.g. 2026-05-10 or "2026-05-10 09:30", local time)',
  flowOptTo: "End of the period to analyze (a date alone means the end of that day)",
  flowOptPageWidth: "Width of each page thumbnail (px)",
  flowOptSlopeWidth: "Width of the band area between two thumbnail columns (px)",
  flowOptTitle: "Title",
  errDate: (name: string, value: string) =>
    `Can't interpret the date for ${name}: ${value} (e.g. 2026-05-10 or "2026-05-10 09:30")`,
  noRevisionsInRange: "No tracked changes in the requested period (--from / --to).",
  flowDone: (sessions: number, pages: number, out: string) =>
    `Wrote the edit flow of ${sessions} sessions (up to ${pages} pages) to ${out}.`,
  flowDefaultTitle: "Edit flow",
  flowSession: (i: number, range: string) => `Session ${i}: ${range}`,
  flowChars: (sign: string, n: string) => `${sign}${n} chars`,
  flowBulkNote: (n: string) => `incl. ${n} chars bulk-inserted`,
  flowMovedNote: (n: string) => `${n} chars moved/reordered`,
  flowLegendFine: "Fine-grained edits (darker = more)",
  flowLegendBulk: "Bulk insertion / replacement (darker = larger share of the paragraph)",
  flowLegendDeleted: "Deleted (mark at a page's right edge = deleted in the next session)",
  flowLegendMoved: "Moved / reordered (right-edge mark = moved in the next session)",
  flowLegendUnchanged: "Unchanged",
  flowLegendNote1:
    "Each column is the document at that point; bands between columns show each paragraph/figure (a table counts " +
    "as one) moving and changing height during that session.",
  flowLegendNote2: "Pages are a schematic (line and figure sizes and page breaks are approximate).",
  flowColumnStart: "Start of session 1",
  flowColumnEnd: (k: number) => `End of session ${k}`,
  flowIdle: (gap: string) => `(${gap} idle)`,
  errNoFlowSessions: "No sessions to draw (sessions is empty).",

  // docx-ai-suspicion-score
  scoreDescription:
    "Analyze a Word file with Track Changes and score, from 0 (no suspicion) to 100 (strong suspicion), how much " +
    "of the text looks like it was written elsewhere and pasted in rather than typed and revised in Word.",
  scoreArgInput: ".docx file to analyze",
  scoreOptOutput: "Write the result to a JSON file (default: stdout)",
  scoreOptMinChars: "Ignore insertions shorter than this many characters (to avoid false positives)",
  scoreOptBurstLow: "Character count at which the burst score is 0",
  scoreOptBurstHigh: "Character count at which the burst score is 100",
  scoreOptRateLow: "Insertion speed (chars/sec) at which the rate score is 0",
  scoreOptRateHigh: "Insertion speed (chars/sec) at which the rate score is 100",
  scoreOptMaxWeight: "Weight of the highest event score in the overall score",
  scoreOptPretty: "Pretty-print the JSON",
  scoreErrorPrefix: "Error:",
  scoreDone: (score: number, risk: string, out: string) =>
    `Done: wrote score=${score} (${risk}) to ${out}. Note: this is a heuristic indication, not evidence of misconduct.`,
};

type Catalog = typeof en;

const ja: Catalog = {
  errorFor: (file, msg) => `エラー (${file}): ${msg}`,
  doneFor: (file, msg) => `完了 (${file}): ${msg}`,
  warning: (msg) => `警告: ${msg}`,
  fileNotFound: (p) => `ファイルが見つかりません: ${p}`,
  errMultiOutput: "エラー: 複数ファイルを指定した場合、-o/--output は使用できません。",
  errPositive: (name) => `${name} には正の数値を指定してください。`,
  optLang: "表示言語 (en / ja)。既定は OS のロケールから判定",
  configLoaded: (p) => `設定ファイルを読み込みました: ${p}`,
  configInvalid: (p, msg) => `設定ファイル ${p} を読み込めません: ${msg}`,
  configNotMapping: "「キー: 値」の形で書いてください",
  configUnknownKey: (key, p) => `${p} の "${key}" は不明なキーのため無視します`,
  configOutputIgnored: (p) => `複数のファイルが指定されたため、${p} の output は使いません`,
  configBadValue: (key, value, p) => `${p} の "${key}" の値 "${value}" は不正なため無視します`,
  optDrop:
    "デスクトップからのドラッグ&ドロップ起動モード。-o未指定時、出力先のファイル名に入力ファイルの最終更新日時を付ける" +
    "(macOSのFinderドロップレットやWindowsエクスプローラーからの直接ドロップ等、" +
    "ターミナルを介さない起動を想定。Windows上ではさらに結果をメッセージボックスで表示する)",
  optPreserveHistory:
    "文書が「保存時に個人情報(変更履歴の作成者・日時)を削除する」設定の場合に、その設定を外し、" +
    "「変更履歴の記録」もオンにして上書き保存する (元のファイルは <名前>.backup-<日時>.docx として残す)。" +
    "Word 等で文書が開かれている場合は書き換えずにエラーにする",
  optPreserveHistoryAlias: "--preserve-history の別名",
  optCheckHistorySettings:
    "出力は作らず、各ファイルの設定を確認して結果を標準出力に書く。" +
    "1行目が ok または needs-fix、needs-fix の場合は2行目以降に確認用の文面 (ドロップレットからの利用を想定)",
  optBulkChars: "同じ作成者・同じ時刻にまとめて挿入された文字数がこれ以上なら一括挿入 (オレンジ) とみなす",
  defaultTitle: (prefix, name, time) => `${prefix}: ${name} (最終更新 ${time})`,
  titleTime: (y, mo, d, h, mi) => `${y}/${mo}/${d} ${h}:${mi}`,

  preservePrompt: (file, problem, action) =>
    `${file}: ${problem}\n--preserve-history の処理を実行し、${action}。` +
    "Word でこの文書を開いている場合は先に閉じてください。\n実行しますか？ [y/N] ",
  preserveDone: (backup) =>
    "変更履歴の作成者と日時を保存し、変更履歴を記録する設定に書き換えて上書き保存しました" +
    ` (元のファイル: ${backup})。以後の編集から日時が記録されます。`,
  trackRevisionsNote: "あわせて「変更履歴の記録」もオンにします。",
  enableHistoryPrompt: (trackNote) =>
    "この文書は編集者の名前と編集日時を保存するための設定が無効化されています。有効にしますか？\n\n" +
    "OK を選ぶと文書の設定を書き換えて上書き保存します (元のファイルはバックアップとして同じフォルダに残します)。" +
    trackNote +
    "Word でこの文書を開いている場合は、先に閉じてください。" +
    "すでに削除された過去の編集日時は復元できません。",
  settingsProblem:
    "この文書は「保存時にファイルのプロパティから個人情報を削除する」設定が有効なため、" +
    "保存時に変更履歴の作成者と日時が削除されます。",
  preserveAction: (trackRevisions) =>
    "この設定を外し" +
    (trackRevisions ? "" : "、「変更履歴の記録」をオンにし") +
    "て上書き保存します (元のファイルはバックアップとして同じフォルダに残します)",
  settingsWarning: (problem, action) => `警告: ${problem}--preserve-history を指定すると、${action}。`,
  errNoWmlNamespace: "settings.xml に WordprocessingML の名前空間が見つかりません。",
  errNoSettingsElement: "settings.xml に w:settings 要素が見つかりません。",
  errSettingsStructure: "settings.xml の構造を解釈できませんでした。",
  errPatchFailed: "settings.xml の書き換えに失敗しました (書き換え後も設定が有効化されていません)。",
  openInWord: (lock) =>
    `Word でこの文書が開かれています (所有者ファイル ${lock} があります)。` +
    "Word で文書を閉じてから再度実行してください。" +
    "Word を閉じてもこのファイルが残っている場合は、異常終了時の残骸なので削除してください。",
  openElsewhere: "他のアプリケーションがこの文書を開いています。閉じてから再度実行してください。",

  undatedCause:
    "Word の「保存時にファイルのプロパティから個人情報を削除する」設定が有効なため、日時が削除されています。" +
    "すでに削除された日時は復元できません。",
  undatedRevisions: (n, cause) =>
    `変更履歴 (w:ins / w:del) は ${n} 件ありますが、すべて日時 (w:date) が記録されていないため時系列解析できません。${cause}`,
  noRevisionsTrackingOff:
    "変更履歴 (w:ins / w:del) が見つかりませんでした。" +
    "このファイルは「変更履歴の記録」がオフの状態で保存されています。" +
    "記録していた場合でも、変更をすべて承諾すると変更履歴は消えます。",
  noRevisionsTrackingOn:
    "変更履歴 (w:ins / w:del) が見つかりませんでした。" +
    "「変更履歴の記録」はオンですが、まだ変更が無いか、変更がすべて承諾/元に戻されています。",
  errNoDocumentXml: "word/document.xml が見つかりません。",
  errNoBody: "word/document.xml に本文 (w:body) が見つかりません。",
  noInsertions: "挿入イベントが見つからなかったため、スコアは算出できません (0を返します)。",
  fewEvents: "挿入イベント数が少ないため (5件未満)、統計的な信頼性は低くなります。参考値として扱ってください。",

  chartTitlePrefix: "編集履歴",
  chartDescription:
    "変更履歴(Track Changes)が有効なWordファイル(.docx)から、時系列の追加/削除文字数と総文字数のSVGチャートを生成します。" +
    "複数ファイルを指定するとまとめて処理します(1件ずつ独立に処理し、失敗しても残りは続行します)。",
  chartArgFiles: "解析対象の .docx ファイル (複数指定可。Windowsでエクスプローラーから複数ファイルをドロップした場合に対応)",
  chartOptOutput: "出力するSVGファイルのパス (既定: 入力と同名の .svg)。複数ファイル指定時は使用不可",
  chartOptBucket: "時間バケットの粒度: auto|second|minute|hour|day、または秒数の数値",
  chartOptGap:
    "更新が連続的に行われた期間とそうでない期間を区別する閾値(時間)。" +
    "指定すると、この閾値を超える無編集期間で区切った期間ごとに個別のグラフを作成し、" +
    "水平に並べて表示する(期間の間には無編集期間の長さを表す間隔を挿入)。",
  chartOptWidth: "SVG幅(px) (未指定時、-p使用時は内容に応じて自動計算)",
  chartOptHeight: "SVG高さ(px)",
  chartOptTitle: "チャートタイトル",
  errGapThreshold: "--gap-threshold (-p) には0以上の数値(時間)を指定してください。",
  chartDoneSessions: (events, sessions, out) =>
    `${events}件のイベントを${sessions}個の期間に分割し、${out} に出力しました。`,
  chartDoneBuckets: (buckets, out) => `${buckets}個の時間バケットに集計し、${out} に出力しました。`,
  kindFine: "細かい編集",
  kindBulk: "一括挿入",
  kindMoved: "移動・並べ替え",
  kindDeleted: "削除",
  totalChars: "総文字数",
  barTooltip: (label, n, when) => `${label} ${n}文字 (${when})`,
  totalTooltip: (n, when) => `総文字数 ${n} (${when})`,
  axisAddedDeleted: "文字数(追加/削除)",
  chartDefaultTitle: "編集履歴 (追加/削除文字数・総文字数)",
  chartDefaultTitleSessions: "編集履歴 (期間ごと・追加/削除文字数・総文字数)",
  chartSplitNote: (hours) => `(${hours}時間以上更新が無い期間で区切って表示)`,
  errNoBuckets: "描画する変更履歴がありません (buckets が空です)。",
  errNoSessions: "描画する変更履歴がありません (sessions が空です)。",

  flowTitlePrefix: "編集フロー",
  flowDescription:
    "変更履歴(Track Changes)付きのWordファイル(.docx)を、連続的に編集が行われた時間区間ごとに分け、" +
    "最初の区間の開始時点と各区間の終了時点の文書を、模式的なページのサムネイルの列として時系列順に左から右へ並べ、" +
    "各区間の終了時点では区間内に細かく編集された段落を緑、まとめて挿入・置き換えられた段落をオレンジで塗り、" +
    "列の間にその区間での段落・図表ごとの削除・置き換え・増加・移動を帯で示した編集フロー(SVG)を1枚にまとめます。",
  flowArgFiles: "解析対象の .docx ファイル (複数指定可)",
  flowOptOutput: "出力するSVGファイルのパス (既定: <入力ファイル名>-flow.svg)。複数ファイル指定時は使用不可",
  flowOptGap: "無編集期間がこの時間を超えたら、別の時間区間に分ける",
  flowOptFrom: '対象期間の開始 (例: 2026-05-10 または "2026-05-10 09:30"。ローカル時刻)',
  flowOptTo: "対象期間の終了 (日付だけの場合はその日の終わりまで)",
  flowOptPageWidth: "ページのサムネイルの幅 (px)",
  flowOptSlopeWidth: "サムネイルの列の間 (段落・図表の変化を示す帯) の幅 (px)",
  flowOptTitle: "図のタイトル",
  errDate: (name, value) => `${name} の日時を解釈できません: ${value} (例: 2026-05-10 または "2026-05-10 09:30")`,
  noRevisionsInRange: "指定された期間 (--from / --to) に変更履歴がありません。",
  flowDone: (sessions, pages, out) =>
    `${sessions}個の時間区間 (最大${pages}ページ) の編集フローを ${out} に出力しました。`,
  flowDefaultTitle: "編集フロー",
  flowSession: (i, range) => `区間${i}: ${range}`,
  flowChars: (sign, n) => `${sign}${n}字`,
  flowBulkNote: (n) => `うち一括挿入 ${n}字`,
  flowMovedNote: (n) => `移動・並べ替え ${n}字`,
  flowLegendFine: "細かい編集 (多いほど濃い)",
  flowLegendBulk: "一括挿入・置き換え (段落に占める割合が大きいほど濃い)",
  flowLegendDeleted: "削除 (ページ右端の印は次の区間で削除)",
  flowLegendMoved: "移動・並べ替え (右端の印は次の区間で移動)",
  flowLegendUnchanged: "変化なし",
  flowLegendNote1:
    "各列はその時点の文書で、列の間の帯は段落・図表 (表は1つ) ごとの、その区間での位置と高さの変化です。",
  flowLegendNote2: "ページは模式図です (行・図の大きさとページ割りは近似)。",
  flowColumnStart: "区間1の開始時点",
  flowColumnEnd: (k) => `区間${k}の終了時点`,
  flowIdle: (gap) => `(${gap} 無編集)`,
  errNoFlowSessions: "描画する時間区間がありません (sessions が空です)。",

  scoreDescription:
    "変更履歴(Track Changes)付きWordファイルを解析し、「WordのGUIでタイプ・校閲せず、" +
    "外部で作文した完成文を貼り付けたのでは」と疑われる不自然な文字数増加の程度を" +
    "0(疑いなし)〜100(疑い濃厚)のスコアとして算出します。",
  scoreArgInput: "解析対象の .docx ファイル",
  scoreOptOutput: "結果をJSONファイルに出力 (省略時は標準出力)",
  scoreOptMinChars: "この文字数未満の挿入は無視 (誤検知防止)",
  scoreOptBurstLow: "バーストスコア0となる文字数境界",
  scoreOptBurstHigh: "バーストスコア100となる文字数境界",
  scoreOptRateLow: "速度スコア0となる挿入速度(文字/秒)",
  scoreOptRateHigh: "速度スコア100となる挿入速度(文字/秒)",
  scoreOptMaxWeight: "文書全体スコアにおける最大イベントスコアの重み",
  scoreOptPretty: "JSONを整形して出力する",
  scoreErrorPrefix: "エラー:",
  scoreDone: (score, risk, out) =>
    `完了: スコア=${score} (${risk}) を ${out} に出力しました。` +
    " ※これはヒューリスティックな参考値であり、不正の証拠にはなりません。",
};

const catalogs: Record<Lang, Catalog> = { en, ja };

type Key = keyof Catalog;
type Args<K extends Key> = Catalog[K] extends (...args: infer A) => string ? A : [];

/** 現在の言語のメッセージを返す */
export function t<K extends Key>(key: K, ...args: Args<K>): string {
  const m = catalogs[getLang()][key] as string | ((...a: unknown[]) => string);
  return typeof m === "function" ? m(...args) : m;
}

/** 表示幅の見積もり (全角12px相当の文字は1、その他は0.55) × フォントサイズ */
export function estimateLabelWidth(text: string, fontSize: number): number {
  let w = 0;
  for (const ch of text) w += ch.codePointAt(0)! >= 0x1100 ? 1 : 0.55;
  return w * fontSize;
}
