#!/usr/bin/env node
/**
 * docx-revision-snapshot: 区切りごとに提出された文書を通し番号付きで保管し、通し解析する。
 *
 *   docx-revision-snapshot <保管先> [提出物のファイル・フォルダ...]
 *
 * 提出物を渡すと、文書ごとに次の番号で保管し (直前の回と同じ内容なら保管しない)、その文書について
 *  - 各回のチャートとフロー (<名前>-s01.svg, <名前>-s01-flow.svg)
 *  - 通し解析のチャート (<名前>-through.svg。1回の提出を1つのパネルにして並べる)
 * を作り直す。各回は前回の提出物と照合して改ざんの痕跡を調べる (第1回は --template があればそれと)。
 * 最後に、保管先の直下に一覧 (index.html と summary.csv) を作る。提出物を渡さなければ、保管済みのすべての文書について
 * 作り直すだけを行う。
 */
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import { DocxRevisionData, describeMissingRevisions, extractRevisions } from "../lib/docxRevisions";
import { buildBuckets } from "../lib/timeBuckets";
import { renderRevisionChart, renderSessionedRevisionChart } from "../lib/svgChart";
import { parseDocxLayout } from "../lib/docxLayout";
import { applyClassification, highlightCategoryMap } from "../lib/classifiers";
import { buildFlow, DEFAULT_FLOW_OPTIONS, FlowResult } from "../lib/flow";
import { renderFlowSvg } from "../lib/flowSvg";
import { buildDefaultTitle } from "../lib/filenames";
import { DEFAULT_BULK_CHARS } from "../lib/insertionKinds";
import type { CategoryRegistry } from "../lib/categories";
import {
  addTamperWarning,
  checkTamperEvidence,
  localize,
  TAMPER_NOTE,
  TAMPERED_SUFFIX,
  TamperEvidence,
  TamperReport,
} from "../lib/tamperEvidence";
import { datedBeforePreviousEvidence, mergeSnapshots, snapshotSessions } from "../lib/snapshots";
import {
  addSnapshot,
  keyBaseName,
  keyDir,
  listSnapshotKeys,
  readManifest,
  snapshotKey,
  SnapshotEntry,
} from "../node/snapshotArchive";
import { analyzeWithRules, langOption, resolveRules } from "./common";
import "../node/locale";
import { getLang, initLangFromArgv, t } from "../lib/i18n";
import { applyToolConfig, loadToolConfig, loadToolConfigOrExit } from "./config";

const TOOL = "docx-revision-snapshot";

const loadedConfig = loadToolConfig(TOOL);
initLangFromArgv(process.argv, loadedConfig.config?.values.lang);
const toolConfig = loadToolConfigOrExit(TOOL, loadedConfig);

/** 入力の .docx と、そのフォルダからの相対パス */
interface InputFile {
  file: string;
  relative: string;
}

function expandSnapshotInputs(inputs: string[]): InputFile[] {
  const out: InputFile[] = [];
  for (const arg of inputs) {
    const abs = path.resolve(arg);
    if (!fs.existsSync(abs)) throw new Error(t("fileNotFound", abs));
    if (!fs.statSync(abs).isDirectory()) {
      out.push({ file: abs, relative: path.basename(abs) });
      continue;
    }
    const walk = (dir: string) => {
      for (const e of fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
        if (e.name.startsWith(".")) continue;
        const p = path.join(dir, e.name);
        if (e.isDirectory()) walk(p);
        else if (e.isFile() && /\.docx$/i.test(e.name) && !e.name.startsWith("~$")) {
          out.push({ file: p, relative: path.relative(abs, p).split(path.sep).join("/") });
        }
      }
    };
    walk(abs);
  }
  return out;
}

/** 痕跡の有無で名前を変える図を書き出し、もう一方の名前の古いファイルを消す */
function writeFigure(dir: string, stem: string, svg: string, tampered: boolean): string {
  const name = `${stem}${tampered ? TAMPERED_SUFFIX : ""}.svg`;
  const other = `${stem}${tampered ? "" : TAMPERED_SUFFIX}.svg`;
  fs.writeFileSync(path.join(dir, name), svg, "utf-8");
  fs.rmSync(path.join(dir, other), { force: true });
  return name;
}

interface SnapshotRow {
  n: number;
  file: string;
  sourceModified: string;
  archivedAt: string;
  newChanges: number;
  carriedChanges: number;
  inserted: number;
  deleted: number;
  evidence: { id: string; message: string }[];
  chart?: string;
  flow?: string;
  note?: string;
}

interface KeyReport {
  key: string;
  through?: string;
  snapshots: SnapshotRow[];
}

/** 1つの文書 (キー) の、各回と通し解析の図を作る */
async function processKey(archive: string, key: string, options: Record<string, any>): Promise<KeyReport> {
  const dir = keyDir(archive, key);
  const manifest = readManifest(dir);
  if (!manifest || manifest.snapshots.length === 0) return { key, snapshots: [] };
  const base = keyBaseName(key);
  const bulkChars = Number(options.bulkChars);
  const gapHours = Number(options.gapThreshold);
  const rules = resolveRules(options, bulkChars);
  const template = options.template ? await fs.promises.readFile(path.resolve(String(options.template))) : undefined;

  const datas: DocxRevisionData[] = [];
  const reports: (TamperReport | undefined)[] = [];
  const rows: SnapshotRow[] = [];
  /** 各回の図の元 (変更履歴の無い回は undefined) */
  const figures: ({ chart: string; flow: FlowResult; categories: CategoryRegistry } | undefined)[] = [];
  let registry: CategoryRegistry | undefined;
  let prev: Buffer | undefined;

  for (const entry of manifest.snapshots) {
    const buf = await fs.promises.readFile(path.join(dir, entry.file));
    const data = await extractRevisions(buf);
    const missing = describeMissingRevisions(data);
    // 前回の提出物を基準にする。教員が承諾のためにロックをかけ直すと salt が変わるため、パスワードは比べない
    const tamper =
      options.tamperCheck === false
        ? undefined
        : await checkTamperEvidence(buf, prev ? { template: prev, compareLockPassword: false } : { template });
    const row: SnapshotRow = {
      n: entry.n,
      file: entry.file,
      sourceModified: entry.sourceModified,
      archivedAt: entry.archivedAt,
      newChanges: 0,
      carriedChanges: 0,
      inserted: 0,
      deleted: 0,
      evidence: [],
      note: missing,
    };
    reports.push(tamper);
    datas.push(data);
    rows.push(row);
    if (missing) {
      figures.push(undefined);
    } else {
      const model = await parseDocxLayout(buf);
      const analysis = await analyzeWithRules(model, rules, gapHours);
      applyClassification(data.events, analysis.classification);
      registry = analysis.categories;
      const title = buildDefaultTitle(t("chartTitlePrefix"), entry.file, new Date(entry.sourceModified));
      const chart = renderRevisionChart(buildBuckets(data.events, data.baselineCharCount, "auto"), {
        eventRange: { start: data.events[0].date, end: data.events[data.events.length - 1].date },
        width: 1100,
        height: 550,
        title,
        categories: analysis.categories,
      });
      const flow = buildFlow(model, { gapThresholdHours: gapHours, bulkChars, highlightOf: highlightCategoryMap(analysis.classification) });
      figures.push({ chart, flow, categories: analysis.categories });
    }
    prev = buf;
  }

  // 各回で新しく加わった変更 (前回までと重複するものを除く)
  const merged = mergeSnapshots(datas);
  merged.deltas.forEach((d, i) => {
    const extra = datedBeforePreviousEvidence(d);
    if (extra && reports[i]) reports[i] = { ...reports[i]!, evidence: [...reports[i]!.evidence, extra] };
    Object.assign(rows[i], {
      newChanges: d.newEvents.length,
      carriedChanges: d.carriedEvents,
      inserted: d.inserted,
      deleted: d.deleted,
      evidence: (reports[i]?.evidence ?? []).map((e) => ({ id: e.id, message: localize(e.message) })),
    });
  });

  // 各回の図
  for (const [i, row] of rows.entries()) {
    const figs = figures[i];
    if (!figs) continue;
    const report = reports[i];
    const tampered = (report?.evidence.length ?? 0) > 0;
    const stem = entryStem(manifest.snapshots[i]);
    row.chart = writeFigure(dir, stem, addTamperWarning(figs.chart, report), tampered);
    if (figs.flow.sessions.length) {
      const flowSvg = renderFlowSvg(figs.flow, {
        title: buildDefaultTitle(t("flowTitlePrefix"), row.file, new Date(row.sourceModified)),
        categories: figs.categories,
      });
      row.flow = writeFigure(dir, `${stem}-flow`, addTamperWarning(flowSvg, report), tampered);
    }
  }

  // 通し解析
  let through: string | undefined;
  const { sessions, indices } = snapshotSessions(merged);
  if (sessions.length) {
    const allEvidence: TamperEvidence[] = [];
    reports.forEach((r, i) => {
      for (const e of r?.evidence ?? []) {
        allEvidence.push({ ...e, message: { en: `#${rows[i].n}: ${localize(e.message, "en")}`, ja: `第${rows[i].n}回: ${localize(e.message, "ja")}` } });
      }
    });
    const svg = renderSessionedRevisionChart(sessions, merged.baselineCharCount, "auto", {
      title: t("snapThroughTitle", base, manifest.snapshots.length),
      categories: registry,
      sessionLabels: indices.map((i) => ({
        text: t("snapPanelLabel", rows[i].n),
        warn: (reports[i]?.evidence.length ?? 0) > 0,
        tooltip: t("snapPanelTooltip", rows[i].file, rows[i].inserted, rows[i].deleted),
      })),
    });
    const report: TamperReport = { checkedAgainstTemplate: true, evidence: allEvidence, note: TAMPER_NOTE };
    through = writeFigure(dir, `${base}-through`, addTamperWarning(svg, report), allEvidence.length > 0);
  }

  const result: KeyReport = { key, through, snapshots: rows };
  fs.writeFileSync(path.join(dir, "report.json"), JSON.stringify(result, null, 2) + "\n", "utf-8");
  return result;
}

function entryStem(e: SnapshotEntry): string {
  return e.file.replace(/\.docx$/i, "");
}

function csvCell(v: unknown): string {
  const s = v === undefined || v === null ? "" : String(v);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
}

/** 保管先の直下に一覧 (summary.csv と index.html) を作る */
function writeIndex(archive: string, reports: KeyReport[]): void {
  const cols = t("snapColumns").split("|");
  const href = (key: string, file?: string) => (file ? encodeURI(`${key}/${file}`) : "");
  const lines = [cols.join(",")];
  for (const r of reports) {
    for (const s of r.snapshots) {
      lines.push(
        [r.key, s.n, s.file, s.sourceModified, s.archivedAt, s.newChanges, s.carriedChanges, s.inserted, s.deleted,
          s.evidence.map((e) => e.id).join(" "), href(r.key, s.chart), href(r.key, s.flow), href(r.key, r.through), s.note]
          .map(csvCell)
          .join(",")
      );
    }
  }
  fs.writeFileSync(path.join(archive, "summary.csv"), "﻿" + lines.join("\r\n") + "\r\n", "utf-8");

  const link = (h: string, label: string) => (h ? `<a href="${escapeHtml(h)}">${escapeHtml(label)}</a>` : "");
  const fmt = (iso: string) => new Date(iso).toLocaleString(getLang() === "ja" ? "ja-JP" : "en-US");
  const sections = reports
    .map((r) => {
      const rows = r.snapshots
        .map(
          (s) =>
            `<tr${s.evidence.length ? ' class="warn"' : ""}><td class="n">${s.n}</td><td>${escapeHtml(s.file)}</td>` +
            `<td>${escapeHtml(fmt(s.sourceModified))}</td><td>${escapeHtml(fmt(s.archivedAt))}</td>` +
            `<td class="n">${s.newChanges}</td><td class="n">${s.carriedChanges}</td><td class="n">${s.inserted}</td><td class="n">${s.deleted}</td>` +
            `<td>${link(href(r.key, s.chart), t("snapChart"))} ${link(href(r.key, s.flow), t("snapFlow"))}</td>` +
            `<td>${escapeHtml([...s.evidence.map((e) => e.message), s.note ?? ""].filter(Boolean).join(" "))}</td></tr>`
        )
        .join("\n");
      return (
        `<h2>${escapeHtml(r.key)}</h2>\n<p>${link(href(r.key, r.through), t("snapThroughLink"))}</p>\n` +
        `<table><thead><tr>${[cols[1], cols[2], cols[3], cols[4], cols[5], cols[6], cols[7], cols[8], `${t("snapChart")} / ${t("snapFlow")}`, cols[9]]
          .map((c) => `<th>${escapeHtml(c)}</th>`)
          .join("")}</tr></thead>\n<tbody>\n${rows}\n</tbody></table>`
      );
    })
    .join("\n");
  const html = `<!doctype html>
<html lang="${getLang()}"><head><meta charset="utf-8"><title>${escapeHtml(t("snapIndexTitle"))}</title>
<style>
body{font:14px/1.5 -apple-system,"Segoe UI","Hiragino Sans","Yu Gothic UI",sans-serif;margin:24px;color:#1d2026}
table{border-collapse:collapse;width:100%;margin-bottom:24px}th,td{border-bottom:1px solid #d9dde3;padding:6px 8px;text-align:left;vertical-align:top}
th{font-size:12px;color:#5f6672}td.n{text-align:right;font-variant-numeric:tabular-nums}tr.warn td{color:#b42318}
p.note{color:#5f6672}
</style></head><body>
<h1>${escapeHtml(t("snapIndexTitle"))}</h1>
<p class="note">${escapeHtml(t("snapIndexNote"))} ${escapeHtml(localize(TAMPER_NOTE))}</p>
${sections}
</body></html>
`;
  fs.writeFileSync(path.join(archive, "index.html"), html, "utf-8");
}

const program = new Command();

program
  .name(TOOL)
  .description(t("snapDescription"))
  .argument("<archive>", t("snapArgArchive"))
  .argument("[inputs...]", t("snapArgInputs"))
  .option("--template <file.docx>", t("snapOptTemplate"))
  .option("--key-depth <n>", t("snapOptKeyDepth"))
  .option("-p, --gap-threshold <hours>", t("flowOptGap"), String(DEFAULT_FLOW_OPTIONS.gapThresholdHours))
  .option("--bulk-chars <n>", t("optBulkChars"), String(DEFAULT_BULK_CHARS))
  .option("--rules <file>", t("optRules"))
  .option("--no-tamper-check", t("optNoTamperCheck"))
  .addOption(langOption())
  .action(async (archiveArg: string, inputs: string[], options) => {
    try {
      const archive = path.resolve(archiveArg);
      fs.mkdirSync(archive, { recursive: true });
      const depth = options.keyDepth !== undefined ? parseInt(options.keyDepth, 10) : undefined;
      if (depth !== undefined && !(depth > 0)) throw new Error(t("errPositive", "--key-depth"));
      if (options.template && !fs.existsSync(path.resolve(options.template))) {
        throw new Error(t("errTemplateNotFound", path.resolve(options.template)));
      }

      // 1. 保管
      const touched = new Set<string>();
      for (const input of expandSnapshotInputs(inputs)) {
        const key = snapshotKey(input.relative, depth);
        const r = await addSnapshot(archive, key, input.file);
        console.error(
          r.added
            ? t("snapAdded", key, r.entry.n, path.join(r.dir, r.entry.file))
            : t("snapUnchanged", key, r.entry.n)
        );
        touched.add(key);
      }
      if (inputs.length > 0 && touched.size === 0) throw new Error(t("errNoDocxFound"));

      // 2. 図 (提出物を渡さなければ、保管済みのすべての文書)
      const keys = inputs.length ? [...touched] : listSnapshotKeys(archive);
      if (keys.length === 0) throw new Error(t("snapNoSnapshots", archive));
      let failed = false;
      for (const key of keys) {
        try {
          const r = await processKey(archive, key, options);
          const warn = r.snapshots.some((s) => s.evidence.length > 0);
          if (warn) console.error(t("warning", t("snapEvidenceFound", key)));
          if (r.through) console.log(path.join(keyDir(archive, key), r.through));
        } catch (err) {
          failed = true;
          console.error(t("errorFor", key, err instanceof Error ? err.message : String(err)));
        }
      }

      // 3. 一覧 (保管済みのすべての文書。作り直さなかった文書は前回の report.json を使う)
      const reports: KeyReport[] = [];
      for (const key of listSnapshotKeys(archive)) {
        const f = path.join(keyDir(archive, key), "report.json");
        if (fs.existsSync(f)) reports.push(JSON.parse(fs.readFileSync(f, "utf-8")) as KeyReport);
      }
      writeIndex(archive, reports);
      console.error(t("snapIndexWritten", path.join(archive, "index.html")));
      if (failed) process.exit(1);
    } catch (err) {
      console.error(t("errorFor", TOOL, err instanceof Error ? err.message : String(err)));
      process.exit(1);
    }
  });

if (toolConfig) applyToolConfig(program, toolConfig);

program.parseAsync(process.argv);
