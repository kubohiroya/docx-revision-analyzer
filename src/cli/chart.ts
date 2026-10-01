#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { describeMissingRevisions, extractRevisionsFromBuffer } from "../lib/docxRevisions";
import { buildBuckets, BucketSpec } from "../lib/timeBuckets";
import { renderRevisionChart, renderSessionedRevisionChart } from "../lib/svgChart";
import { splitIntoSessions } from "../lib/sessions";
import { buildDefaultTitle, buildDropOutputPath } from "../lib/filenames";
import { classifyInsertionsByLevels, DEFAULT_BULK_CHARS } from "../lib/insertionKinds";
import { parseDocxLayout } from "../lib/docxLayout";
import {
  addAnalysisOptions,
  addCommonOptions,
  analyzeWithRules,
  checkAndFixHistorySettings,
  FileResult,
  runForFiles,
  resolveRules,
  writeAnalysisJson,
} from "./common";
import { initLangFromArgv, t } from "../lib/i18n";
import { applyToolConfig, loadToolConfig, loadToolConfigOrExit } from "./config";

// 実行ファイルと同じフォルダの設定ファイルを読む (lang の指定があれば表示言語にも使う)
const loadedConfig = loadToolConfig("docx-revision-chart");
// ヘルプの文言はコマンドの定義時に決まるため、先に表示言語を確定する
initLangFromArgv(process.argv, loadedConfig.config?.values.lang);
const toolConfig = loadToolConfigOrExit("docx-revision-chart", loadedConfig);

const program = new Command();

async function processOne(
  inputFile: string,
  options: Record<string, any>,
  explicitOutput: string | undefined
): Promise<FileResult> {
  const resolved = path.resolve(inputFile);
  if (!fs.existsSync(resolved)) {
    return { input: inputFile, ok: false, message: t("fileNotFound", resolved) };
  }

  const warnings: string[] = await checkAndFixHistorySettings(resolved, options, "docx-revision-chart");
  const buf = await fs.promises.readFile(resolved);
  const data = await extractRevisionsFromBuffer(buf);
  const missing = describeMissingRevisions(data);
  if (missing) {
    // 時系列解析できるイベントが無い場合は SVG を作らずにエラーとする
    return { input: inputFile, ok: false, notes: warnings, message: missing };
  }

  const bulkChars = parseFloat(options.bulkChars);
  if (!Number.isFinite(bulkChars) || bulkChars <= 0) {
    return { input: inputFile, ok: false, notes: warnings, message: t("errPositive", "--bulk-chars") };
  }
  const rules = resolveRules(options, bulkChars);
  const analysis = analyzeWithRules(await parseDocxLayout(buf), rules);
  classifyInsertionsByLevels(data.events, analysis.levelOf);
  const rulesNote = options.rules ? t("rulesNote", rules.ruleSet) : undefined;

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
  const writeJson = async (svgOut: string) => {
    const jsonOut = writeAnalysisJson(options, svgOut, "docx-revision-chart", inputFile, analysis);
    if (jsonOut) warnings.push(t("jsonWritten", jsonOut));
  };

  if (options.gapThreshold !== undefined) {
    const thresholdHours = parseFloat(options.gapThreshold);
    if (Number.isNaN(thresholdHours) || thresholdHours < 0) {
      return {
        input: inputFile,
        ok: false,
        message: t("errGapThreshold"),
      };
    }

    const sessions = splitIntoSessions(data.events, thresholdHours);
    const svg = renderSessionedRevisionChart(sessions, data.baselineCharCount, bucketSpec, {
      width,
      height: parseInt(options.height, 10),
      title: options.title ?? buildDefaultTitle(t("chartTitlePrefix"), inputFile, fs.statSync(resolved).mtime),
      gapThresholdHours: thresholdHours,
      note: rulesNote,
    });
    fs.writeFileSync(outFile, svg, "utf-8");
    await writeJson(outFile);

    return {
      input: inputFile,
      ok: true,
      outFile,
      message:
        t("chartDoneSessions", data.events.length, sessions.length, outFile),
      notes: warnings,
    };
  } else {
    const buckets = buildBuckets(data.events, data.baselineCharCount, bucketSpec);
    const svg = renderRevisionChart(buckets, {
      eventRange: { start: data.events[0].date, end: data.events[data.events.length - 1].date },
      width: width ?? 1100,
      height: parseInt(options.height, 10),
      title: options.title ?? buildDefaultTitle(t("chartTitlePrefix"), inputFile, fs.statSync(resolved).mtime),
      note: rulesNote,
    });
    fs.writeFileSync(outFile, svg, "utf-8");
    await writeJson(outFile);

    return {
      input: inputFile,
      ok: true,
      outFile,
      notes: warnings,
      message: t("chartDoneBuckets", buckets.length, outFile),
    };
  }
}

program
  .name("docx-revision-chart")
  .description(t("chartDescription"))
  .argument("<files...>", t("chartArgFiles"))
  .option("-o, --output <file.svg>", t("chartOptOutput"))
  .option("-b, --bucket <spec>", t("chartOptBucket"), "auto")
  .option("-p, --gap-threshold <hours>", t("chartOptGap"))
  .option("--bulk-chars <n>", t("optBulkChars"), String(DEFAULT_BULK_CHARS))
  .option("-w, --width <number>", t("chartOptWidth"))
  .option("-H, --height <number>", t("chartOptHeight"), "550")
  .option("-t, --title <text>", t("chartOptTitle"))
  .action(async (files: string[], options) => {
    await runForFiles("docx-revision-chart", files, options, (file, output) =>
      processOne(file, options, output)
    );
  });

addAnalysisOptions(program);
addCommonOptions(program);

if (toolConfig) applyToolConfig(program, toolConfig);

program.parseAsync(process.argv);
