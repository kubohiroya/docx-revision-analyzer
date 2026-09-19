#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { spawnSync } from "child_process";
import { extractRevisionsFromFile } from "../lib/docxRevisions";
import { buildBuckets, BucketSpec } from "../lib/timeBuckets";
import { renderRevisionChart, renderSessionedRevisionChart } from "../lib/svgChart";
import { splitIntoSessions } from "../lib/sessions";
import { buildDropOutputPath } from "../lib/filenames";

const program = new Command();

interface FileResult {
  input: string;
  ok: boolean;
  outFile?: string;
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

async function processOne(
  inputFile: string,
  options: Record<string, any>,
  explicitOutput: string | undefined
): Promise<FileResult> {
  const resolved = path.resolve(inputFile);
  if (!fs.existsSync(resolved)) {
    return { input: inputFile, ok: false, message: `ファイルが見つかりません: ${resolved}` };
  }

  const data = await extractRevisionsFromFile(resolved);
  const warnings: string[] = [];
  if (data.events.length === 0) {
    warnings.push(
      "警告: 変更履歴 (w:ins / w:del) が見つかりませんでした。" +
        "このファイルは「変更履歴の記録」を有効にして編集されたものか確認してください。"
    );
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
        [...warnings, `${data.events.length}件のイベントを${sessions.length}個の期間に分割し、${outFile} に出力しました。`].join(
          " "
        ),
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
      message: [...warnings, `${buckets.length}個の時間バケットに集計し、${outFile} に出力しました。`].join(" "),
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
  .action(async (files: string[], options) => {
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
      if (r.ok) {
        console.error(`完了 (${r.input}): ${r.message}`);
        console.log(r.outFile);
      } else {
        console.error(`エラー (${r.input}): ${r.message}`);
      }
    }

    if (options.drop) {
      const summary = results
        .map((r) => (r.ok ? `✓ ${r.outFile}` : `✗ ${r.input}: ${r.message}`))
        .join("\n");
      showWindowsMessageBox("docx-revision-chart", summary);
    }

    if (results.some((r) => !r.ok)) {
      process.exit(1);
    }
  });

program.parseAsync(process.argv);
