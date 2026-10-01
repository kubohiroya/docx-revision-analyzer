#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { describeMissingRevisions, extractRevisionsFromFile } from "../lib/docxRevisions";
import { parseDocxLayout } from "../lib/docxLayout";
import { extractRevisionPositions } from "../lib/revisionPositions";
import { detectInsertionWindows } from "../lib/insertionWindows";
import { buildFlow, DEFAULT_FLOW_OPTIONS } from "../lib/flow";
import { renderFlowSvg } from "../lib/flowSvg";
import { buildDefaultTitle, buildDropOutputPath } from "../lib/filenames";
import {
  addAnalysisOptions,
  addCommonOptions,
  checkAndFixHistorySettings,
  FileResult,
  runForFiles,
  windowOptionsFrom,
  writeAnalysisJson,
} from "./common";
import { initLangFromArgv, t } from "../lib/i18n";
import { applyToolConfig, loadToolConfig, loadToolConfigOrExit } from "./config";

const TOOL = "docx-revision-flow";

// 実行ファイルと同じフォルダの設定ファイルを読む (lang の指定があれば表示言語にも使う)
const loadedConfig = loadToolConfig(TOOL);
// ヘルプの文言はコマンドの定義時に決まるため、先に表示言語を確定する
initLangFromArgv(process.argv, loadedConfig.config?.values.lang);
const toolConfig = loadToolConfigOrExit(TOOL, loadedConfig);

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
  const parsed = Date.parse(value);
  if (Number.isNaN(parsed)) {
    throw new Error(t("errDate", name, value));
  }
  return new Date(parsed);
}

function positiveNumber(value: string, name: string): number {
  const n = parseFloat(value);
  if (!Number.isFinite(n) || n <= 0) throw new Error(t("errPositive", name));
  return n;
}

async function processOne(
  inputFile: string,
  options: Record<string, any>,
  explicitOutput: string | undefined
): Promise<FileResult> {
  const resolved = path.resolve(inputFile);
  if (!fs.existsSync(resolved)) {
    return { input: inputFile, ok: false, message: t("fileNotFound", resolved) };
  }

  const gapThresholdHours = positiveNumber(options.gapThreshold, "--gap-threshold (-p)");
  const bulkChars = positiveNumber(options.bulkChars, "--bulk-chars");
  const pageWidth = positiveNumber(options.pageWidth, "--page-width");
  const slopeWidth = positiveNumber(options.slopeWidth, "--slope-width");
  const from = options.from ? parseDateOption(options.from, "--from", false) : undefined;
  const to = options.to ? parseDateOption(options.to, "--to", true) : undefined;
  const windowOptions = windowOptionsFrom(options);

  const notes = await checkAndFixHistorySettings(resolved, options, TOOL);

  const model = await parseDocxLayout(await fs.promises.readFile(resolved));
  const result = buildFlow(model, { gapThresholdHours, bulkChars, from, to });

  if (result.sessions.length === 0) {
    // 変更履歴そのものが無い (または日時が無い) のか、期間の指定で外れたのかを区別して伝える
    const missing = describeMissingRevisions(await extractRevisionsFromFile(resolved));
    return {
      input: inputFile,
      ok: false,
      notes,
      message: missing ?? t("noRevisionsInRange"),
    };
  }

  let outFile: string;
  if (explicitOutput) {
    outFile = explicitOutput;
  } else {
    const base = resolved.replace(/\.docx$/i, "");
    outFile = options.drop
      ? buildDropOutputPath(`${base}-flow.docx`, fs.statSync(resolved).mtime, ".svg")
      : `${base}-flow.svg`;
  }

  const svg = renderFlowSvg(result, {
    title: options.title ?? buildDefaultTitle(t("flowTitlePrefix"), inputFile, fs.statSync(resolved).mtime),
    pageWidth,
    slopeWidth,
  });
  fs.writeFileSync(outFile, svg, "utf-8");
  const jsonOut = writeAnalysisJson(options, outFile, TOOL, inputFile, detectInsertionWindows(extractRevisionPositions(model), windowOptions));
  if (jsonOut) notes.push(t("jsonWritten", jsonOut));

  const pages = Math.max(...result.sessions.flatMap((s) => [s.startPages.length, s.endPages.length]));
  return {
    input: inputFile,
    ok: true,
    outFile,
    notes,
    message: t("flowDone", result.sessions.length, pages, outFile),
  };
}

program
  .name(TOOL)
  .description(t("flowDescription"))
  .argument("<files...>", t("flowArgFiles"))
  .option("-o, --output <file.svg>", t("flowOptOutput"))
  .option("-p, --gap-threshold <hours>", t("flowOptGap"), String(DEFAULT_FLOW_OPTIONS.gapThresholdHours))
  .option("--from <datetime>", t("flowOptFrom"))
  .option("--to <datetime>", t("flowOptTo"))
  .option("--bulk-chars <n>", t("optBulkChars"), String(DEFAULT_FLOW_OPTIONS.bulkChars))
  .option("--page-width <px>", t("flowOptPageWidth"), "150")
  .option("--slope-width <px>", t("flowOptSlopeWidth"), "72")
  .option("-t, --title <text>", t("flowOptTitle"))
  .action(async (files: string[], options) => {
    await runForFiles(TOOL, files, options, (file, output) => processOne(file, options, output));
  });

addAnalysisOptions(program);
addCommonOptions(program);

if (toolConfig) applyToolConfig(program, toolConfig);

program.parseAsync(process.argv);
