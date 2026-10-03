#!/usr/bin/env node
/**
 * docx-revision-versions: 変更履歴の日時から、過去のある時点の文書 (版) を復元して書き出す。
 *
 *   docx-revision-versions <files...> [--at <日時>]... [--every <間隔>] [--sessions]
 *
 * 版は <出力先>/<名前>-<YYYYMMDD-HHMMSS>.docx (または .txt) に書き出し、各版の時刻を
 * <出力先>/<名前>.versions.json にまとめる。出力先の既定は入力と同じフォルダの <名前>-versions/。
 */
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { describeMissingRevisions, extractRevisions } from "../lib/docxRevisions";
import { DEFAULT_FLOW_OPTIONS } from "../lib/flow";
import { formatTimestampForFilename } from "../lib/filenames";
import { docxPlainText, openVersionSource, versionPoints } from "../lib/versions";
import { FileResult, langOption, parseDateOption, runForFiles } from "./common";
import "../node/locale";
import { initLangFromArgv, t } from "../lib/i18n";
import { applyToolConfig, loadToolConfig, loadToolConfigOrExit } from "./config";

const TOOL = "docx-revision-versions";

const loadedConfig = loadToolConfig(TOOL);
initLangFromArgv(process.argv, loadedConfig.config?.values.lang);
const toolConfig = loadToolConfigOrExit(TOOL, loadedConfig);

const FORMATS = ["docx", "txt"] as const;
type Format = (typeof FORMATS)[number];

/** "30s" / "10m" / "2h" / "1d" または秒数を、ミリ秒に変換する */
function parseInterval(value: string): number {
  const m = value.trim().match(/^(\d+(?:\.\d+)?)\s*([smhd]?)$/i);
  const unit = { "": 1, s: 1, m: 60, h: 3600, d: 86400 }[(m?.[2] ?? "").toLowerCase()];
  const ms = m && unit ? Number(m[1]) * unit * 1000 : NaN;
  if (!(ms >= 1000)) throw new Error(t("errInterval", "--every", value));
  return ms;
}

function parseFormats(value: string): Format[] {
  const list = value.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  const bad = list.find((f) => !(FORMATS as readonly string[]).includes(f));
  if (bad !== undefined || list.length === 0) throw new Error(t("errVersionFormat", value));
  return [...new Set(list)] as Format[];
}

async function processOne(inputFile: string, options: Record<string, any>): Promise<FileResult> {
  const resolved = path.resolve(inputFile);
  if (!fs.existsSync(resolved)) {
    return { input: inputFile, ok: false, message: t("fileNotFound", resolved) };
  }
  const formats = parseFormats(options.format);
  // 設定ファイルの at は1つの文字列で入る
  const at = ([] as string[]).concat(options.at ?? []).map((v) => parseDateOption(v, "--at", true));
  const everyMs = options.every !== undefined ? parseInterval(options.every) : undefined;
  const gap = parseFloat(options.gapThreshold);
  if (!(gap > 0)) throw new Error(t("errPositive", "--gap-threshold (-p)"));
  // 何も指定しなければ、時間区間ごとの版を作る
  const sessions = options.sessions || (at.length === 0 && everyMs === undefined);

  const buf = await fs.promises.readFile(resolved);
  const source = await openVersionSource(buf);
  if (at.length === 0 && source.revisionDates.length === 0) {
    // 変更履歴そのものが無いのか、日時が無いのかを伝える
    const missing = describeMissingRevisions(await extractRevisions(buf)) ?? t("noRevisionsInRange");
    return { input: inputFile, ok: false, message: missing };
  }
  const points = versionPoints(source.revisionDates, { at, everyMs, sessionGapHours: sessions ? gap : undefined });

  const base = path.basename(resolved).replace(/\.docx$/i, "");
  const outDir = options.outDir
    ? path.resolve(options.outDir)
    : path.join(path.dirname(resolved), `${base}-versions`);
  fs.mkdirSync(outDir, { recursive: true });

  const used = new Set<string>();
  const manifest = [];
  for (const p of points) {
    // 同じ秒の版はファイル名に番号を付けて区別する
    let stem = `${base}-${formatTimestampForFilename(p.at)}`;
    for (let n = 2; used.has(stem); n++) stem = `${base}-${formatTimestampForFilename(p.at)}-${n}`;
    used.add(stem);
    const docx = await source.docxAt(p.at.getTime());
    const files: string[] = [];
    if (formats.includes("docx")) {
      fs.writeFileSync(path.join(outDir, `${stem}.docx`), docx);
      files.push(`${stem}.docx`);
    }
    if (formats.includes("txt")) {
      fs.writeFileSync(path.join(outDir, `${stem}.txt`), await docxPlainText(docx), "utf-8");
      files.push(`${stem}.txt`);
    }
    manifest.push({
      at: p.at.toISOString(),
      label: p.label,
      changes: source.revisionDates.filter((d) => d.getTime() <= p.at.getTime()).length,
      files,
    });
  }
  const manifestFile = path.join(outDir, `${base}.versions.json`);
  fs.writeFileSync(
    manifestFile,
    JSON.stringify(
      {
        tool: TOOL,
        input: path.basename(resolved),
        totalChanges: source.revisionDates.length,
        undatedChanges: source.undatedCount,
        versions: manifest,
      },
      null,
      2
    ) + "\n",
    "utf-8"
  );

  const notes: string[] = [];
  if (source.undatedCount > 0) notes.push(t("versionsUndated", source.undatedCount));
  return {
    input: inputFile,
    ok: true,
    outFile: outDir,
    notes,
    message: t("versionsDone", points.length, outDir),
  };
}

const collect = (value: string, previous: string[] = []) => [...previous, value];

const program = new Command();
program
  .name(TOOL)
  .description(t("versionsDescription"))
  .argument("<files...>", t("versionsArgFiles"))
  .option("--at <datetime>", t("versionsOptAt"), collect)
  .option("--every <interval>", t("versionsOptEvery"))
  .option("--sessions", t("versionsOptSessions"))
  .option("-p, --gap-threshold <hours>", t("flowOptGap"), String(DEFAULT_FLOW_OPTIONS.gapThresholdHours))
  .option("-f, --format <list>", t("versionsOptFormat"), "docx")
  .option("-o, --out-dir <dir>", t("versionsOptOutDir"))
  .addOption(langOption())
  .action(async (files: string[], options) => {
    await runForFiles(TOOL, files, options, (file) => processOne(file, options));
  });

if (toolConfig) applyToolConfig(program, toolConfig);

program.parseAsync(process.argv);
