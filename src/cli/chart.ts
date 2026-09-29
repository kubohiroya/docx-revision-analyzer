#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { describeMissingRevisions, extractRevisionsFromFile } from "../lib/docxRevisions";
import { buildBuckets, BucketSpec } from "../lib/timeBuckets";
import { renderRevisionChart, renderSessionedRevisionChart } from "../lib/svgChart";
import { splitIntoSessions } from "../lib/sessions";
import { buildDropOutputPath } from "../lib/filenames";
import { classifyInsertions, DEFAULT_BULK_CHARS } from "../lib/insertionKinds";
import { addCommonOptions, checkAndFixHistorySettings, FileResult, runForFiles } from "./common";

const program = new Command();

async function processOne(
  inputFile: string,
  options: Record<string, any>,
  explicitOutput: string | undefined
): Promise<FileResult> {
  const resolved = path.resolve(inputFile);
  if (!fs.existsSync(resolved)) {
    return { input: inputFile, ok: false, message: `ファイルが見つかりません: ${resolved}` };
  }

  const warnings: string[] = await checkAndFixHistorySettings(resolved, options, "docx-revision-chart");
  const data = await extractRevisionsFromFile(resolved);
  const missing = describeMissingRevisions(data);
  if (missing) {
    // 時系列解析できるイベントが無い場合は SVG を作らずにエラーとする
    return { input: inputFile, ok: false, notes: warnings, message: missing };
  }

  const bulkChars = parseFloat(options.bulkChars);
  if (!Number.isFinite(bulkChars) || bulkChars <= 0) {
    return { input: inputFile, ok: false, notes: warnings, message: "--bulk-chars には正の数値を指定してください。" };
  }
  classifyInsertions(data.events, bulkChars);

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
  .option(
    "--bulk-chars <n>",
    "同じ作成者・同じ時刻にまとめて挿入された文字数がこれ以上なら一括挿入 (オレンジ) とみなす。" +
      "それ以外の挿入は細かい編集 (緑)、削除された文章と同じ内容の挿入は移動・並べ替え (青) として積み上げる",
    String(DEFAULT_BULK_CHARS)
  )
  .option("-w, --width <number>", "SVG幅(px) (未指定時、-p使用時は内容に応じて自動計算)")
  .option("-H, --height <number>", "SVG高さ(px)", "550")
  .option("-t, --title <text>", "チャートタイトル")
  .action(async (files: string[], options) => {
    await runForFiles("docx-revision-chart", files, options, (file, output) =>
      processOne(file, options, output)
    );
  });

addCommonOptions(program);

program.parseAsync(process.argv);
