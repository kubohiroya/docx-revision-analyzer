#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { extractRevisionsFromFile } from "../lib/docxRevisions";
import {
  computeSuspicionScore,
  DEFAULT_SUSPICION_OPTIONS,
} from "../lib/suspicionScore";

const program = new Command();

program
  .name("docx-ai-suspicion-score")
  .description(
    "変更履歴(Track Changes)付きWordファイルを解析し、「WordのGUIでタイプ・校閲せず、" +
      "外部で作文した完成文を貼り付けたのでは」と疑われる不自然な文字数増加の程度を" +
      "0(疑いなし)〜100(疑い濃厚)のスコアとして算出します。"
  )
  .argument("<input.docx>", "解析対象の .docx ファイル")
  .option("-o, --output <file.json>", "結果をJSONファイルに出力 (省略時は標準出力)")
  .option(
    "--min-chars <n>",
    "この文字数未満の挿入は無視 (誤検知防止)",
    String(DEFAULT_SUSPICION_OPTIONS.minCharsToFlag)
  )
  .option(
    "--burst-low <n>",
    "バーストスコア0となる文字数境界",
    String(DEFAULT_SUSPICION_OPTIONS.burstLowChars)
  )
  .option(
    "--burst-high <n>",
    "バーストスコア100となる文字数境界",
    String(DEFAULT_SUSPICION_OPTIONS.burstHighChars)
  )
  .option(
    "--rate-low <cps>",
    "速度スコア0となる挿入速度(文字/秒)",
    String(DEFAULT_SUSPICION_OPTIONS.rateLowCps)
  )
  .option(
    "--rate-high <cps>",
    "速度スコア100となる挿入速度(文字/秒)",
    String(DEFAULT_SUSPICION_OPTIONS.rateHighCps)
  )
  .option(
    "--max-weight <0-1>",
    "文書全体スコアにおける最大イベントスコアの重み",
    String(DEFAULT_SUSPICION_OPTIONS.maxWeight)
  )
  .option("--pretty", "JSONを整形して出力する", false)
  .action(async (inputFile: string, options) => {
    try {
      const resolved = path.resolve(inputFile);
      if (!fs.existsSync(resolved)) {
        console.error(`エラー: ファイルが見つかりません: ${resolved}`);
        process.exit(1);
      }

      const data = await extractRevisionsFromFile(resolved);

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
          `完了: スコア=${report.score} (${report.riskLevel}) を ${options.output} に出力しました。` +
            " ※これはヒューリスティックな参考値であり、不正の証拠にはなりません。"
        );
      } else {
        console.log(json);
      }
    } catch (err) {
      console.error("エラー:", err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

program.parseAsync(process.argv);
