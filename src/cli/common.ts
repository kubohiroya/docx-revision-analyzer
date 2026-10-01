/**
 * CLI 共通の処理: 変更履歴の設定チェック (--preserve-history)、確認ダイアログ、
 * 複数ファイルの処理と結果の表示。
 */
import { Command, Option } from "commander";
import * as fs from "fs";
import * as path from "path";
import { spawnSync } from "child_process";
import * as readline from "readline";
import JSZip from "jszip";
import {
  buildEnableHistoryPrompt,
  describeHistorySettingsProblem,
  describeHistorySettingsWarning,
  describePreserveHistoryAction,
  enableHistoryPreservation,
  needsHistoryFix,
  parseHistorySettings,
} from "../lib/historySettings";
import { t } from "../lib/i18n";
import { parse as parseYaml } from "yaml";
import type { DocxLayoutModel } from "../lib/docxLayout";
import { extractRevisionPositions } from "../lib/revisionPositions";
import {
  detectInsertionWindows,
  InsertionWindowResult,
  insertionWindowsToJson,
  WindowOptions,
} from "../lib/insertionWindows";
import { CATEGORY_IDS, CategoryRegistry, contrastWithWhite, MIN_GRAPHIC_CONTRAST } from "../lib/categories";
import {
  assignLevels,
  categoriesFromRules,
  defaultRuleSet,
  levelsByInsertion,
  parseRuleSet,
  RuleLevel,
  RuleSet,
  RuleSetError,
} from "../lib/insertionRules";

export interface FileResult {
  input: string;
  ok: boolean;
  outFile?: string;
  /** 結果の前に伝える補足 (設定の書き換え報告や警告)。1件ずつ改行して表示する */
  notes?: string[];
  message: string;
}

/**
 * Windows専用: --drop モードで結果をコンソールなしでも確認できるよう、
 * ネイティブのメッセージボックスを表示する (ベストエフォート、失敗しても無視)。
 * macOS側の通知はFinderドロップレット(AppleScript)側で表示するため、
 * ここではWindowsのみを対象にする。
 */
export function showWindowsMessageBox(title: string, message: string): void {
  if (process.platform !== "win32") return;
  try {
    const psTitle = title.replace(/'/g, "''");
    const psMessage = message.replace(/'/g, "''");
    const script =
      `Add-Type -AssemblyName System.Windows.Forms; ` +
      `[System.Windows.Forms.MessageBox]::Show('${psMessage}', '${psTitle}') | Out-Null`;
    spawnSync("powershell", ["-NoProfile", "-NonInteractive", "-Command", script], {
      stdio: "ignore",
      windowsHide: false,
    });
  } catch {
    // ベストエフォート: 通知に失敗しても処理全体は成功として扱う
  }
}

/**
 * Windows専用: OK/キャンセルの確認ダイアログを表示し、OK が押されたら true を返す。
 * 表示に失敗した場合は false (変更しない側) として扱う。
 */
function askWindowsOkCancel(title: string, message: string): boolean {
  if (process.platform !== "win32") return false;
  try {
    const psTitle = title.replace(/'/g, "''");
    const psMessage = message.replace(/'/g, "''");
    const script =
      `Add-Type -AssemblyName System.Windows.Forms; ` +
      `[System.Windows.Forms.MessageBox]::Show('${psMessage}', '${psTitle}', 'OKCancel', 'Question')`;
    const r = spawnSync("powershell", ["-NoProfile", "-NonInteractive", "-Command", script], {
      encoding: "utf-8",
      windowsHide: false,
    });
    return r.stdout.trim() === "OK";
  } catch {
    return false;
  }
}

/**
 * ターミナルで y/N を尋ねる (質問は標準エラー出力に書き、標準出力は結果のパス専用に保つ)。
 * y / yes 以外 (空入力を含む) は No とみなす。
 */
function askYesNo(question: string): Promise<boolean> {
  const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  return new Promise((resolve) => {
    let answered = false;
    rl.question(question, (answer) => {
      // rl.close() は close イベントを同期的に発火するため、先に結果を確定させる
      answered = true;
      resolve(/^y(es)?$/i.test(answer.trim()));
      rl.close();
    });
    // 回答前の Ctrl+D (入力終了) は No として扱う
    rl.on("close", () => {
      if (answered) return;
      process.stderr.write("\n");
      resolve(false);
    });
  });
}

/** 対話的に y/N を尋ねられる状況か (ドロップ起動やパイプ経由ではない) */
function canAskInteractively(options: Record<string, any>): boolean {
  return !options.drop && Boolean(process.stdin.isTTY) && Boolean(process.stderr.isTTY);
}

async function readSettingsOf(file: string) {
  const zip = await JSZip.loadAsync(await fs.promises.readFile(file));
  const settingsFile = zip.file("word/settings.xml");
  return parseHistorySettings(settingsFile ? await settingsFile.async("string") : undefined);
}

/**
 * 解析の前に、文書が変更履歴の作成者・日時を保存する設定になっているかを確認する。
 * --preserve-history 指定時 (または Windows の --drop で OK が押された時) は設定を書き換える。
 * 返り値は利用者に伝えるメッセージ。
 */
export async function checkAndFixHistorySettings(
  resolved: string,
  options: Record<string, any>,
  toolName: string
): Promise<string[]> {
  const settings = await readSettingsOf(resolved);
  if (!needsHistoryFix(settings)) return [];

  let fix = Boolean(options.preserveHistory);
  if (!fix && options.drop && process.platform === "win32") {
    fix = askWindowsOkCancel(
      `${toolName} - ${path.basename(resolved)}`,
      buildEnableHistoryPrompt(settings)!
    );
  } else if (!fix && canAskInteractively(options)) {
    fix = await askYesNo(
      t(
        "preservePrompt",
        path.basename(resolved),
        describeHistorySettingsProblem(settings)!,
        describePreserveHistoryAction(settings)
      )
    );
  }
  if (!fix) return [describeHistorySettingsWarning(settings)!];

  const r = await enableHistoryPreservation(resolved);
  return [t("preserveDone", r.backupPath ?? "")];
}

/** --drop / --preserve-history / --check-history-settings / --lang を登録する */
export function addCommonOptions(program: Command): Command {
  return program
    .option("--drop", t("optDrop"))
    .option("--preserve-history", t("optPreserveHistory"))
    .addOption(new Option("--preserveHistory", t("optPreserveHistoryAlias")).hideHelp())
    .option("--check-history-settings", t("optCheckHistorySettings"))
    .addOption(langOption());
}

/** 判定ルール (--rules)、解析結果の JSON 出力 (--json)、挿入の窓 (--window-*) のオプションを登録する */
export function addAnalysisOptions(program: Command): Command {
  return program
    .option("--rules <file>", t("optRules"))
    .option("--json [file.json]", t("optJson"))
    .option("--window-seconds <s>", t("optWindowSeconds"))
    .option("--window-chars <n>", t("optWindowChars"))
    .option("--window-paras <n>", t("optWindowParas"));
}

function nonNegativeOption(value: unknown, name: string): number {
  const n = parseFloat(String(value));
  if (!Number.isFinite(n) || n < 0) throw new Error(t("errNonNegative", name));
  return n;
}

/** ルールファイルを読み込んで検証する。不正ならエラー */
export function loadRuleSetFile(file: string): RuleSet {
  let raw: unknown;
  try {
    raw = parseYaml(fs.readFileSync(file, "utf-8"));
  } catch (err) {
    throw new Error(t("rulesInvalid", file, err instanceof Error ? err.message : String(err)));
  }
  try {
    return parseRuleSet(raw);
  } catch (err) {
    if (err instanceof RuleSetError) throw new Error(t("rulesInvalid", file, err.message));
    throw err;
  }
}

/**
 * 使う判定ルールと窓の設定を決める。--rules が無ければ --bulk-chars から作る既定ルール。
 * --window-* を指定した場合は、ルールの窓の設定のうちその項目を置き換える。
 */
export function resolveRules(options: Record<string, any>, bulkChars: number): RuleSet {
  const rules = options.rules ? loadRuleSetFile(String(options.rules)) : defaultRuleSet(bulkChars);
  const window: WindowOptions = { ...rules.window };
  if (options.windowSeconds !== undefined) window.seconds = nonNegativeOption(options.windowSeconds, "--window-seconds");
  if (options.windowChars !== undefined) window.chars = nonNegativeOption(options.windowChars, "--window-chars");
  if (options.windowParas !== undefined) window.paras = nonNegativeOption(options.windowParas, "--window-paras");
  return { ...rules, window };
}

export interface RuleAnalysis {
  rules: RuleSet;
  /** 図の色のカテゴリ (既定のカテゴリ + ルールのレベル。レベルが既定の一括挿入を置き換える) */
  categories: CategoryRegistry;
  /** ルールについての注意 (色のコントラストが低いなど) */
  warnings: string[];
  windows: InsertionWindowResult;
  /** 窓ごとのレベル (windows.windows と同じ順) */
  levels: (RuleLevel | undefined)[];
  /** 挿入の w:id → レベル */
  levelOf: Map<string, RuleLevel>;
}

/** レイアウトモデルから挿入の窓を求め、判定ルールでレベルを付ける */
export function analyzeWithRules(model: DocxLayoutModel, rules: RuleSet): RuleAnalysis {
  const windows = detectInsertionWindows(extractRevisionPositions(model), rules.window);
  const levels = assignLevels(rules, windows);
  const categories = new CategoryRegistry();
  categories.unregister(CATEGORY_IDS.bulk);
  for (const c of categoriesFromRules(rules)) categories.register(c);
  const warnings = rules.levels
    .filter((lv) => contrastWithWhite(lv.color) < MIN_GRAPHIC_CONTRAST)
    .map((lv) => t("warning", t("rulesLowContrast", lv.id, contrastWithWhite(lv.color).toFixed(1))));
  return { rules, categories, warnings, windows, levels, levelOf: levelsByInsertion(windows, levels) };
}

/**
 * --json が指定されていれば、解析結果の JSON を書き出してそのパスを返す。
 * ファイル名が省略された場合 (設定ファイルの json: true を含む) は、SVG と同じ名前の .json にする。
 */
export function writeAnalysisJson(
  options: Record<string, any>,
  svgOut: string,
  tool: string,
  inputFile: string,
  analysis: RuleAnalysis
): string | undefined {
  const json = options.json;
  if (json === undefined || json === false || json === "false") return undefined;
  const outFile = json === true || json === "true" ? svgOut.replace(/\.svg$/i, "") + ".json" : String(json);
  const w = insertionWindowsToJson(analysis.windows);
  const body = {
    tool,
    input: path.basename(inputFile),
    ruleSet: analysis.rules.ruleSet,
    levels: analysis.rules.levels.map((lv) => ({ id: lv.id, label: lv.label, color: lv.color, when: lv.when })),
    window: w.window,
    timeResolutionSec: w.timeResolutionSec,
    windows: w.windows.map((x, i) => ({ ...x, level: analysis.levels[i]?.id ?? null })),
  };
  fs.writeFileSync(outFile, JSON.stringify(body, null, 2) + "\n", "utf-8");
  return outFile;
}

/** --lang <en|ja>。値は initLangFromArgv が先に読むため、ここではヘルプと引数の検証のために登録する */
export function langOption(): Option {
  return new Option("--lang <lang>", t("optLang")).choices(["en", "ja"]);
}

/**
 * 各ファイルを processOne で処理し、結果を表示して終了コードを決める。
 * --check-history-settings 指定時は設定の確認だけを行う。
 */
export async function runForFiles(
  toolName: string,
  files: string[],
  options: Record<string, any>,
  processOne: (file: string, explicitOutput: string | undefined) => Promise<FileResult>
): Promise<void> {
  if (options.checkHistorySettings) {
    let failed = false;
    for (const file of files) {
      try {
        const settings = await readSettingsOf(path.resolve(file));
        const prompt = buildEnableHistoryPrompt(settings);
        console.log(prompt ? `needs-fix\n${prompt}` : "ok");
      } catch (err) {
        failed = true;
        console.error(t("errorFor", file, err instanceof Error ? err.message : String(err)));
      }
    }
    process.exit(failed ? 1 : 0);
  }

  if (options.output && files.length > 1) {
    console.error(t("errMultiOutput"));
    process.exit(1);
  }

  const results: FileResult[] = [];
  for (const file of files) {
    try {
      results.push(await processOne(file, files.length === 1 ? options.output : undefined));
    } catch (err) {
      results.push({
        input: file,
        ok: false,
        message: err instanceof Error ? err.message : String(err),
      });
    }
  }

  for (const r of results) {
    for (const note of r.notes ?? []) console.error(`${r.input}: ${note}`);
    if (r.ok) {
      console.error(t("doneFor", r.input, r.message));
      console.log(r.outFile);
    } else {
      console.error(t("errorFor", r.input, r.message));
    }
  }

  if (options.drop) {
    const summary = results
      .map((r) =>
        [...(r.notes ?? []), r.ok ? `✓ ${r.outFile}` : `✗ ${r.input}: ${r.message}`].join("\n")
      )
      .join("\n\n");
    showWindowsMessageBox(toolName, summary);
  }

  if (results.some((r) => !r.ok)) {
    process.exit(1);
  }
}
