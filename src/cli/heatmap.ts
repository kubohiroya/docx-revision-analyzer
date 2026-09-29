#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { describeMissingRevisions, extractRevisionsFromFile } from "../lib/docxRevisions";
import { parseDocxLayout } from "../lib/docxLayout";
import { buildHeatmap, DEFAULT_HEATMAP_OPTIONS } from "../lib/heatmap";
import { renderHeatmapSvg } from "../lib/heatmapSvg";
import { buildDropOutputPath } from "../lib/filenames";
import { addCommonOptions, checkAndFixHistorySettings, FileResult, runForFiles } from "./common";

const TOOL = "docx-revision-heatmap";
const program = new Command();

/**
 * "YYYY-MM-DD" / "YYYY-MM-DD HH:mm" (ローカル時刻) または ISO 8601 形式の日時を解釈する。
 * 日付だけの場合、endOfDay なら その日の終わり (23:59:59.999) とする。
 */
function parseDateOption(value: string, name: string, endOfDay: boolean): Date {
  const m = value.trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?$/);
  if (m) {
    const [, y, mo, d, h, mi] = m;
    if (h === undefined) {
      return endOfDay
        ? new Date(Number(y), Number(mo) - 1, Number(d), 23, 59, 59, 999)
        : new Date(Number(y), Number(mo) - 1, Number(d));
    }
    return new Date(Number(y), Number(mo) - 1, Number(d), Number(h), Number(mi), endOfDay ? 59 : 0, endOfDay ? 999 : 0);
  }
  const t = Date.parse(value);
  if (Number.isNaN(t)) {
    throw new Error(`${name} の日時を解釈できません: ${value} (例: 2026-05-10 または "2026-05-10 09:30")`);
  }
  return new Date(t);
}

function positiveNumber(value: string, name: string): number {
  const n = parseFloat(value);
  if (!Number.isFinite(n) || n <= 0) throw new Error(`${name} には正の数値を指定してください。`);
  return n;
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

  const gapThresholdHours = positiveNumber(options.gapThreshold, "--gap-threshold (-p)");
  const bulkChars = positiveNumber(options.bulkChars, "--bulk-chars");
  const pageWidth = positiveNumber(options.pageWidth, "--page-width");
  const slopeWidth = positiveNumber(options.slopeWidth, "--slope-width");
  const from = options.from ? parseDateOption(options.from, "--from", false) : undefined;
  const to = options.to ? parseDateOption(options.to, "--to", true) : undefined;

  const notes = await checkAndFixHistorySettings(resolved, options, TOOL);

  const model = await parseDocxLayout(await fs.promises.readFile(resolved));
  const result = buildHeatmap(model, { gapThresholdHours, bulkChars, from, to });

  if (result.sessions.length === 0) {
    // 変更履歴そのものが無い (または日時が無い) のか、期間の指定で外れたのかを区別して伝える
    const missing = describeMissingRevisions(await extractRevisionsFromFile(resolved));
    return {
      input: inputFile,
      ok: false,
      notes,
      message: missing ?? "指定された期間 (--from / --to) に変更履歴がありません。",
    };
  }

  let outFile: string;
  if (explicitOutput) {
    outFile = explicitOutput;
  } else {
    const base = resolved.replace(/\.docx$/i, "");
    outFile = options.drop
      ? buildDropOutputPath(`${base}-heatmap.docx`, fs.statSync(resolved).mtime, ".svg")
      : `${base}-heatmap.svg`;
  }

  const svg = renderHeatmapSvg(result, {
    title: options.title ?? `編集ヒートマップ: ${path.basename(inputFile)}`,
    pageWidth,
    slopeWidth,
  });
  fs.writeFileSync(outFile, svg, "utf-8");

  const pages = Math.max(...result.sessions.flatMap((s) => [s.startPages.length, s.endPages.length]));
  return {
    input: inputFile,
    ok: true,
    outFile,
    notes,
    message: `${result.sessions.length}個の時間区間 (最大${pages}ページ) のヒートマップを ${outFile} に出力しました。`,
  };
}

program
  .name(TOOL)
  .description(
    "変更履歴(Track Changes)付きのWordファイル(.docx)を、連続的に編集が行われた時間区間ごとに分け、" +
      "最初の区間の開始時点と各区間の終了時点の文書を、模式的なページのサムネイルの列として時系列順に左から右へ並べ、" +
      "各区間の終了時点では区間内に細かく編集された段落を緑、まとめて挿入・置き換えられた段落をオレンジで塗り、" +
      "列の間にその区間での段落・図表ごとの削除・置き換え・増加を帯で示したヒートマップ(SVG)を1枚にまとめます。"
  )
  .argument("<files...>", "解析対象の .docx ファイル (複数指定可)")
  .option("-o, --output <file.svg>", "出力するSVGファイルのパス (既定: <入力ファイル名>-heatmap.svg)。複数ファイル指定時は使用不可")
  .option(
    "-p, --gap-threshold <hours>",
    "無編集期間がこの時間を超えたら、別の時間区間に分ける",
    String(DEFAULT_HEATMAP_OPTIONS.gapThresholdHours)
  )
  .option("--from <datetime>", "対象期間の開始 (例: 2026-05-10 または \"2026-05-10 09:30\"。ローカル時刻)")
  .option("--to <datetime>", "対象期間の終了 (日付だけの場合はその日の終わりまで)")
  .option(
    "--bulk-chars <n>",
    "同じ時刻にまとめて挿入された文字数がこれ以上なら一括挿入 (オレンジ) とみなす",
    String(DEFAULT_HEATMAP_OPTIONS.bulkChars)
  )
  .option("--page-width <px>", "ページのサムネイルの幅 (px)", "150")
  .option("--slope-width <px>", "開始時点と終了時点のサムネイルの間 (段落・図表の変化を示す帯) の幅 (px)", "72")
  .option("-t, --title <text>", "図のタイトル")
  .action(async (files: string[], options) => {
    await runForFiles(TOOL, files, options, (file, output) => processOne(file, options, output));
  });

addCommonOptions(program);

program.parseAsync(process.argv);
