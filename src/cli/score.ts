#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { describeMissingRevisions, extractRevisionsFromFile } from "../lib/docxRevisions";
import {
  computeSuspicionScore,
  DEFAULT_SUSPICION_OPTIONS,
} from "../lib/suspicionScore";

import { initLangFromArgv, t } from "../lib/i18n";
import { applyToolConfig, loadToolConfig, loadToolConfigOrExit } from "./config";
import { langOption } from "./common";

// 実行ファイルと同じフォルダの設定ファイルを読む (lang の指定があれば表示言語にも使う)
const loadedConfig = loadToolConfig("docx-ai-suspicion-score");
// ヘルプの文言はコマンドの定義時に決まるため、先に表示言語を確定する
initLangFromArgv(process.argv, loadedConfig.config?.values.lang);
const toolConfig = loadToolConfigOrExit("docx-ai-suspicion-score", loadedConfig);

const program = new Command();

program
  .name("docx-ai-suspicion-score")
  .description(t("scoreDescription"))
  .argument("<input.docx>", t("scoreArgInput"))
  .option("-o, --output <file.json>", t("scoreOptOutput"))
  .option("--min-chars <n>", t("scoreOptMinChars"), String(DEFAULT_SUSPICION_OPTIONS.minCharsToFlag))
  .option("--burst-low <n>", t("scoreOptBurstLow"), String(DEFAULT_SUSPICION_OPTIONS.burstLowChars))
  .option("--burst-high <n>", t("scoreOptBurstHigh"), String(DEFAULT_SUSPICION_OPTIONS.burstHighChars))
  .option("--rate-low <cps>", t("scoreOptRateLow"), String(DEFAULT_SUSPICION_OPTIONS.rateLowCps))
  .option("--rate-high <cps>", t("scoreOptRateHigh"), String(DEFAULT_SUSPICION_OPTIONS.rateHighCps))
  .option("--max-weight <0-1>", t("scoreOptMaxWeight"), String(DEFAULT_SUSPICION_OPTIONS.maxWeight))
  .option("--pretty", t("scoreOptPretty"), false)
  .addOption(langOption())
  .action(async (inputFile: string, options) => {
    try {
      const resolved = path.resolve(inputFile);
      if (!fs.existsSync(resolved)) {
        console.error(`${t("scoreErrorPrefix")} ${t("fileNotFound", resolved)}`);
        process.exit(1);
      }

      const data = await extractRevisionsFromFile(resolved);
      const missing = describeMissingRevisions(data);
      if (missing) console.error(t("warning", missing));

      const report = computeSuspicionScore(data.events, {
        minCharsToFlag: parseFloat(options.minChars),
        burstLowChars: parseFloat(options.burstLow),
        burstHighChars: parseFloat(options.burstHigh),
        rateLowCps: parseFloat(options.rateLow),
        rateHighCps: parseFloat(options.rateHigh),
        maxWeight: parseFloat(options.maxWeight),
      });

      const output = {
        file: path.basename(inputFile),
        ...report,
      };

      const json = options.pretty
        ? JSON.stringify(output, null, 2)
        : JSON.stringify(output);

      if (options.output) {
        fs.writeFileSync(options.output, json, "utf-8");
        console.error(
          t("scoreDone", report.score, report.riskLevel, options.output)
        );
      } else {
        console.log(json);
      }
    } catch (err) {
      console.error(t("scoreErrorPrefix"), err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

if (toolConfig) applyToolConfig(program, toolConfig);

program.parseAsync(process.argv);
