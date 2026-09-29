#!/usr/bin/env node
import { Command, Option } from "commander";
import * as fs from "fs";
import * as path from "path";
import { spawnSync } from "child_process";
import * as readline from "readline";
import { describeMissingRevisions, extractRevisionsFromFile } from "../lib/docxRevisions";
import { buildBuckets, BucketSpec } from "../lib/timeBuckets";
import { renderRevisionChart, renderSessionedRevisionChart } from "../lib/svgChart";
import { splitIntoSessions } from "../lib/sessions";
import { buildDropOutputPath } from "../lib/filenames";
import {
  buildEnableHistoryPrompt,
  describeHistorySettingsProblem,
  describeHistorySettingsWarning,
  describePreserveHistoryAction,
  enableHistoryPreservation,
  needsHistoryFix,
  parseHistorySettings,
} from "../lib/historySettings";
import JSZip from "jszip";

const program = new Command();

interface FileResult {
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
function showWindowsMessageBox(title: string, message: string): void {
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
async function checkAndFixHistorySettings(
  resolved: string,
  options: Record<string, any>
): Promise<string[]> {
  const settings = await readSettingsOf(resolved);
  if (!needsHistoryFix(settings)) return [];

  let fix = Boolean(options.preserveHistory);
  if (!fix && options.drop && process.platform === "win32") {
    fix = askWindowsOkCancel(
      `docx-revision-chart - ${path.basename(resolved)}`,
      buildEnableHistoryPrompt(settings)!
    );
  } else if (!fix && canAskInteractively(options)) {
    fix = await askYesNo(
      `${path.basename(resolved)}: ${describeHistorySettingsProblem(settings)}\n` +
        `--preserve-history の処理を実行し、${describePreserveHistoryAction(settings)}。` +
        "Word でこの文書を開いている場合は先に閉じてください。\n" +
        "実行しますか？ [y/N] "
    );
  }
  if (!fix) return [describeHistorySettingsWarning(settings)!];

  const r = await enableHistoryPreservation(resolved);
  return [
    "変更履歴の作成者と日時を保存し、変更履歴を記録する設定に書き換えて上書き保存しました" +
      ` (元のファイル: ${r.backupPath})。以後の編集から日時が記録されます。`,
  ];
}

async function processOne(
  inputFile: string,
  options: Record<string, any>,
  explicitOutput: string | undefined
): Promise<FileResult> {
  const resolved = path.resolve(inputFile);
  if (!fs.existsSync(resolved)) {
    return { input: inputFile, ok: false, message: `ファイルが見つかりません: ${resolved}` };
  }

  const warnings: string[] = await checkAndFixHistorySettings(resolved, options);
  const data = await extractRevisionsFromFile(resolved);
  const missing = describeMissingRevisions(data);
  if (missing) {
    // 時系列解析できるイベントが無い場合は SVG を作らずにエラーとする
    return { input: inputFile, ok: false, notes: warnings, message: missing };
  }

  let bucketSpec: BucketSpec = options.bucket;
  if (/^\d+$/.test(options.bucket)) {
    bucketSpec = parseInt(options.bucket, 10);
  }

  let outFile: string;
  if (explicitOutput) {
    outFile = explicitOutput;
  } else if (options.drop) {
    const mtime = fs.statSync(resolved).mtime;
    outFile = buildDropOutputPath(resolved, mtime, ".svg");
  } else {
    outFile = resolved.replace(/\.docx$/i, "") + ".svg";
  }
  const width = options.width ? parseInt(options.width, 10) : undefined;

  if (options.gapThreshold !== undefined) {
    const thresholdHours = parseFloat(options.gapThreshold);
    if (Number.isNaN(thresholdHours) || thresholdHours < 0) {
      return {
        input: inputFile,
        ok: false,
        message: "--gap-threshold (-p) には0以上の数値(時間)を指定してください。",
      };
    }

    const sessions = splitIntoSessions(data.events, thresholdHours);
    const svg = renderSessionedRevisionChart(sessions, data.baselineCharCount, bucketSpec, {
      width,
      height: parseInt(options.height, 10),
      title: options.title ?? `編集履歴: ${path.basename(inputFile)}`,
      gapThresholdHours: thresholdHours,
    });
    fs.writeFileSync(outFile, svg, "utf-8");

    return {
      input: inputFile,
      ok: true,
      outFile,
      message:
        `${data.events.length}件のイベントを${sessions.length}個の期間に分割し、${outFile} に出力しました。`,
      notes: warnings,
    };
  } else {
    const buckets = buildBuckets(data.events, data.baselineCharCount, bucketSpec);
    const svg = renderRevisionChart(buckets, {
      width: width ?? 1100,
      height: parseInt(options.height, 10),
      title: options.title ?? `編集履歴: ${path.basename(inputFile)}`,
    });
    fs.writeFileSync(outFile, svg, "utf-8");

    return {
      input: inputFile,
      ok: true,
      outFile,
      notes: warnings,
      message: `${buckets.length}個の時間バケットに集計し、${outFile} に出力しました。`,
    };
  }
}

program
  .name("docx-revision-chart")
  .description(
    "変更履歴(Track Changes)が有効なWordファイル(.docx)から、時系列の追加/削除文字数と総文字数のSVGチャートを生成します。" +
      "複数ファイルを指定するとまとめて処理します(1件ずつ独立に処理し、失敗しても残りは続行します)。"
  )
  .argument("<files...>", "解析対象の .docx ファイル (複数指定可。Windowsでエクスプローラーから複数ファイルをドロップした場合に対応)")
  .option("-o, --output <file.svg>", "出力するSVGファイルのパス (既定: 入力と同名の .svg)。複数ファイル指定時は使用不可")
  .option(
    "-b, --bucket <spec>",
    "時間バケットの粒度: auto|second|minute|hour|day、または秒数の数値",
    "auto"
  )
  .option(
    "-p, --gap-threshold <hours>",
    "更新が連続的に行われた期間とそうでない期間を区別する閾値(時間)。" +
      "指定すると、この閾値を超える無編集期間で区切った期間ごとに個別のグラフを作成し、" +
      "水平に並べて表示する(期間の間には無編集期間の長さを表す間隔を挿入)。"
  )
  .option("-w, --width <number>", "SVG幅(px) (未指定時、-p使用時は内容に応じて自動計算)")
  .option("-H, --height <number>", "SVG高さ(px)", "550")
  .option("-t, --title <text>", "チャートタイトル")
  .option(
    "--drop",
    "デスクトップからのドラッグ&ドロップ起動モード。-o未指定時、出力先を" +
      "「<入力と同じディレクトリ>/<ファイル名>-<入力ファイルの最終更新日時>.svg」にする" +
      "(macOSのFinderドロップレットやWindowsエクスプローラーからの直接ドロップ等、" +
      "ターミナルを介さない起動を想定。Windows上ではさらに結果をメッセージボックスで表示する)"
  )
  .option(
    "--preserve-history",
    "文書が「保存時に個人情報(変更履歴の作成者・日時)を削除する」設定の場合に、その設定を外し、" +
      "「変更履歴の記録」もオンにして上書き保存する (元のファイルは <名前>.backup-<日時>.docx として残す)。" +
      "Word 等で文書が開かれている場合は書き換えずにエラーにする"
  )
  .addOption(new Option("--preserveHistory", "--preserve-history の別名").hideHelp())
  .option(
    "--check-history-settings",
    "チャートは作らず、各ファイルの設定を確認して結果を標準出力に書く。" +
      "1行目が ok または needs-fix、needs-fix の場合は2行目以降に確認用の文面 (ドロップレットからの利用を想定)"
  )
  .action(async (files: string[], options) => {
    if (options.checkHistorySettings) {
      let failed = false;
      for (const file of files) {
        try {
          const settings = await readSettingsOf(path.resolve(file));
          const prompt = buildEnableHistoryPrompt(settings);
          console.log(prompt ? `needs-fix\n${prompt}` : "ok");
        } catch (err) {
          failed = true;
          console.error(`エラー (${file}): ${err instanceof Error ? err.message : err}`);
        }
      }
      process.exit(failed ? 1 : 0);
    }

    if (options.output && files.length > 1) {
      console.error("エラー: 複数ファイルを指定した場合、-o/--output は使用できません。");
      process.exit(1);
    }

    const results: FileResult[] = [];
    for (const file of files) {
      try {
        const result = await processOne(file, options, files.length === 1 ? options.output : undefined);
        results.push(result);
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
        console.error(`完了 (${r.input}): ${r.message}`);
        console.log(r.outFile);
      } else {
        console.error(`エラー (${r.input}): ${r.message}`);
      }
    }

    if (options.drop) {
      const summary = results
        .map((r) =>
          [...(r.notes ?? []), r.ok ? `✓ ${r.outFile}` : `✗ ${r.input}: ${r.message}`].join("\n")
        )
        .join("\n\n");
      showWindowsMessageBox("docx-revision-chart", summary);
    }

    if (results.some((r) => !r.ok)) {
      process.exit(1);
    }
  });

program.parseAsync(process.argv);
