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
  needsHistoryFix,
  parseHistorySettings,
} from "../lib/historySettings";
import { t } from "../lib/i18n";
import { enableHistoryPreservation } from "../node/historyFile";
import { lockTemplateFile } from "../node/lockFile";
import { hasLockPassword } from "../lib/trackLock";
import { parse as parseYaml } from "yaml";
import type { DocxLayoutModel } from "../lib/docxLayout";
import { extractRevisionPositions } from "../lib/revisionPositions";
import {
  InsertionWindowResult,
  insertionWindowsToJson,
  WindowOptions,
} from "../lib/insertionWindows";
import type { IntegrityReport } from "../lib/integrity";
import {
  checkTamperEvidence,
  hasTamperEvidence,
  tamperedFileName,
  TamperReport,
} from "../lib/tamperEvidence";
import { FigureAnnotation, FigureTarget, ResolvedAnnotation, resolveAnnotations } from "../lib/figureTargets";
import { CategoryRegistry, contrastWithWhite, MIN_GRAPHIC_CONTRAST } from "../lib/categories";
import {
  ClassificationResult,
  createAnalysisContext,
  defaultClassifiers,
  highlightToJson,
  runClassifiers,
} from "../lib/classifiers";
import {
  assignLevels,
  defaultRuleSet,
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

/**
 * ターミナルで、入力した文字を表示せずに1行読む (パスワード用)。質問は標準エラー出力に書く。
 * Ctrl+C で中止する
 */
function askHidden(question: string): Promise<string> {
  return new Promise((resolve) => {
    const stdin = process.stdin;
    process.stderr.write(question);
    stdin.setRawMode(true);
    stdin.resume();
    stdin.setEncoding("utf8");
    let value = "";
    const finish = () => {
      stdin.setRawMode(false);
      stdin.pause();
      stdin.removeListener("data", onData);
      process.stderr.write("\n");
      resolve(value);
    };
    const onData = (chunk: string) => {
      for (const c of chunk) {
        if (c === "\r" || c === "\n" || c === "\u0004") return finish();
        if (c === "\u0003") {
          stdin.setRawMode(false);
          process.stderr.write("\n");
          process.exit(130);
        }
        if (c === "\u007f" || c === "\b") value = value.slice(0, -1);
        else value += c;
      }
    };
    stdin.on("data", onData);
  });
}

/**
 * --lock で使うパスワードを決める。ターミナルなら2回尋ね、そうでなければ環境変数 DOCX_LOCK_PASSWORD を使う。
 * 空ならパスワード無し (警告を返す)。2回の入力が違えば undefined。
 * chosen は、利用者がパスワード (空を含む) を明示的に決めたか (ターミナルで入力した・環境変数で指定した)
 */
async function lockPassword(
  options: Record<string, any>
): Promise<{ password: string; chosen: boolean; notes: string[] } | undefined> {
  let password = process.env.DOCX_LOCK_PASSWORD ?? "";
  let chosen = process.env.DOCX_LOCK_PASSWORD !== undefined;
  if (!chosen && canAskInteractively(options) && typeof process.stdin.setRawMode === "function") {
    password = await askHidden(t("lockPasswordPrompt"));
    if (password && (await askHidden(t("lockPasswordConfirm"))) !== password) return undefined;
    chosen = true;
  }
  return { password, chosen, notes: password ? [] : [t("warning", t("lockNoPassword"))] };
}

/** --lock: 各ファイルに変更履歴のロックをかけて上書きする */
async function lockFiles(files: string[], options: Record<string, any>): Promise<FileResult[]> {
  const pw = await lockPassword(options);
  if (!pw) {
    console.error(t("lockPasswordMismatch"));
    process.exit(1);
  }
  for (const n of pw.notes) console.error(n);
  const results: FileResult[] = [];
  for (const file of files) {
    try {
      const resolved = path.resolve(file);
      if (!fs.existsSync(resolved)) throw new Error(t("fileNotFound", resolved));
      // パスワードを尋ねられなかったときに、パスワード付きのロックをパスワード無しで置き換えない
      if (!pw.password && !pw.chosen && (await hasLockPassword(await fs.promises.readFile(resolved)))) {
        throw new Error(t("lockKeepsPassword"));
      }
      const r = await lockTemplateFile(resolved, pw.password || undefined);
      const notes: string[] = [];
      if (r.wasLocked) notes.push(t("lockReplaced"));
      if (!r.wasTracking) notes.push(t("lockTrackingTurnedOn"));
      if (r.removedPersonalInfoSetting) notes.push(t("lockPersonalInfoRemoved"));
      results.push({ input: file, ok: true, outFile: resolved, notes, message: t("lockDone", r.backupPath) });
    } catch (err) {
      results.push({ input: file, ok: false, message: err instanceof Error ? err.message : String(err) });
    }
  }
  return results;
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

/** --drop / --preserve-history / --check-history-settings / --lock / --lang を登録する */
export function addCommonOptions(program: Command): Command {
  return program
    .option("--drop", t("optDrop"))
    .option("--preserve-history", t("optPreserveHistory"))
    .addOption(new Option("--preserveHistory", t("optPreserveHistoryAlias")).hideHelp())
    .option("--check-history-settings", t("optCheckHistorySettings"))
    .option("--lock", t("optLock"))
    .addOption(langOption());
}

/** 判定ルール (--rules)、解析結果の JSON 出力 (--json)、挿入の窓 (--window-*) のオプションを登録する */
export function addAnalysisOptions(program: Command): Command {
  return program
    .option("--rules <file>", t("optRules"))
    .option("--annotations <file>", t("optAnnotations"))
    .option("--json [file.json]", t("optJson"))
    .option("--window-seconds <s>", t("optWindowSeconds"))
    .option("--window-chars <n>", t("optWindowChars"))
    .option("--window-paras <n>", t("optWindowParas"))
    .option("--template <file.docx>", t("optTemplate"))
    .option("--no-tamper-check", t("optNoTamperCheck"));
}

/** 読み込んだテンプレート (フォルダを処理するときに毎回読まないよう、パスごとに保持する) */
const templateCache = new Map<string, Buffer>();

/**
 * 改ざんの痕跡を調べる。--no-tamper-check のときは undefined。
 * --template があれば、そのテンプレートとも照合する
 */
export async function checkTamper(buf: Buffer, options: Record<string, any>): Promise<TamperReport | undefined> {
  if (options.tamperCheck === false) return undefined;
  let template: Buffer | undefined;
  if (options.template) {
    const file = path.resolve(String(options.template));
    if (!templateCache.has(file)) {
      if (!fs.existsSync(file)) throw new Error(t("errTemplateNotFound", file));
      templateCache.set(file, await fs.promises.readFile(file));
    }
    template = templateCache.get(file);
  }
  return checkTamperEvidence(buf, { template });
}

/**
 * 痕跡が見つかったときの出力先とメッセージ。-o で出力先を指定した場合は名前を変えない
 */
export function tamperedOutput(
  outFile: string,
  explicitOutput: string | undefined,
  report: TamperReport | undefined
): { outFile: string; note?: string } {
  if (!hasTamperEvidence(report)) return { outFile };
  const ids = report!.evidence.map((e) => e.id).join(", ");
  return { outFile: explicitOutput ? outFile : tamperedFileName(outFile), note: t("warning", t("tamperFound", ids)) };
}

function nonNegativeOption(value: unknown, name: string): number {
  const n = parseFloat(String(value));
  if (!Number.isFinite(n) || n < 0) throw new Error(t("errNonNegative", name));
  return n;
}

/** 注釈ファイル (YAML / JSON) を読む。--annotations が無ければ undefined */
export function loadAnnotations(options: Record<string, any>): Map<string, ResolvedAnnotation> | undefined {
  if (!options.annotations) return undefined;
  const file = String(options.annotations);
  let raw: unknown;
  try {
    raw = parseYaml(fs.readFileSync(file, "utf-8"));
  } catch (err) {
    throw new Error(t("annotationsInvalid", file, err instanceof Error ? err.message : String(err)));
  }
  const list = Array.isArray(raw) ? raw : (raw as { annotations?: unknown })?.annotations;
  if (!Array.isArray(list)) throw new Error(t("annotationsInvalid", file, t("annotationsNotList")));
  return resolveAnnotations(list as FigureAnnotation[]);
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
  /** 図の色のカテゴリ (既定のカテゴリ + 分類器のカテゴリ。ルールのレベルが既定の一括挿入を置き換える) */
  categories: CategoryRegistry;
  /** ルールについての注意 (色のコントラストが低いなど) */
  warnings: string[];
  /** ルールの窓の設定で求めた挿入の窓 */
  windows: InsertionWindowResult;
  /** 窓ごとのレベル (windows.windows と同じ順) */
  levels: (RuleLevel | undefined)[];
  /** 分類器のパイプラインの結果 */
  classification: ClassificationResult;
}

/** レイアウトモデルから解析結果を作り、既定の分類器 (並べ替え・判定ルール) で分類する */
export async function analyzeWithRules(
  model: DocxLayoutModel,
  rules: RuleSet,
  gapThresholdHours: number
): Promise<RuleAnalysis> {
  const ctx = createAnalysisContext(model, extractRevisionPositions(model), gapThresholdHours);
  const classification = await runClassifiers(defaultClassifiers(rules), ctx);
  const windows = ctx.windowsFor(rules.window);
  const warnings = rules.levels
    .filter((lv) => contrastWithWhite(lv.color) < MIN_GRAPHIC_CONTRAST)
    .map((lv) => t("warning", t("rulesLowContrast", lv.id, contrastWithWhite(lv.color).toFixed(1))));
  return {
    rules,
    categories: classification.categories,
    warnings,
    windows,
    levels: assignLevels(rules, windows),
    classification,
  };
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
  analysis: RuleAnalysis,
  integrity?: IntegrityReport,
  figureTargets?: FigureTarget[],
  tamper?: TamperReport
): string | undefined {
  const json = options.json;
  if (json === undefined || json === false || json === "false") return undefined;
  const outFile = json === true || json === "true" ? svgOut.replace(/\.svg$/i, "") + ".json" : String(json);
  const w = insertionWindowsToJson(analysis.windows);
  const body = {
    tool,
    input: path.basename(inputFile),
    ruleSet: analysis.rules.ruleSet,
    classifiers: analysis.classification.classifiers,
    levels: analysis.rules.levels.map((lv) => ({ id: lv.id, label: lv.label, color: lv.color, when: lv.when })),
    window: w.window,
    timeResolutionSec: w.timeResolutionSec,
    windows: w.windows.map((x, i) => ({ ...x, level: analysis.levels[i]?.id ?? null })),
    highlights: analysis.classification.highlights.map(highlightToJson),
    // 整合性の簡易チェック (判定ではなく情報として)
    integrity: integrity ?? null,
    // 改ざんの痕跡 (--no-tamper-check のときは null)
    tamperEvidence: tamper ?? null,
    // 図の部分の一覧 (--annotations の target に使うキー)
    figureTargets: figureTargets ?? [],
  };
  fs.writeFileSync(outFile, JSON.stringify(body, null, 2) + "\n", "utf-8");
  return outFile;
}

/** --lang <en|ja>。値は initLangFromArgv が先に読むため、ここではヘルプと引数の検証のために登録する */
export function langOption(): Option {
  return new Option("--lang <lang>", t("optLang")).choices(["en", "ja"]);
}

/**
 * 引数のフォルダを、その下の .docx (サブフォルダを含む。Word の一時ファイル ~$xxx.docx と隠しファイルは除く) に展開する。
 * ファイルはそのまま残す
 */
export function expandInputs(inputs: string[]): string[] {
  const out: string[] = [];
  const walk = (dir: string) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
      if (e.name.startsWith(".")) continue;
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else if (e.isFile() && /\.docx$/i.test(e.name) && !e.name.startsWith("~$")) out.push(p);
    }
  };
  for (const f of inputs) {
    let isDir = false;
    try {
      isDir = fs.statSync(f).isDirectory();
    } catch {
      isDir = false;
    }
    if (isDir) walk(f);
    else out.push(f);
  }
  return out;
}

/**
 * 各ファイルを processOne で処理し、結果を表示して終了コードを決める。
 * フォルダを渡した場合は、その下の .docx をすべて処理する (それぞれの隣に結果を書き出す)。
 * --check-history-settings 指定時は設定の確認だけを、--lock 指定時はロックをかけるだけを行う。
 */
export async function runForFiles(
  toolName: string,
  files: string[],
  options: Record<string, any>,
  processOne: (file: string, explicitOutput: string | undefined) => Promise<FileResult>
): Promise<void> {
  files = expandInputs(files);
  if (files.length === 0) {
    console.error(t("errNoDocxFound"));
    process.exit(1);
  }
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

  const results: FileResult[] = options.lock ? await lockFiles(files, options) : [];
  for (const file of options.lock ? [] : files) {
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
