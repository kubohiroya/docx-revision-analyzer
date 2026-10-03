/**
 * フォルダの一括処理 (レンダラ側)。
 *
 *  - ローカルのフォルダ: その下の .docx それぞれの隣に、チャート (<名前>.svg) とフロー (<名前>-flow.svg) を作る
 *  - SharePoint / OneDrive のフォルダ (Teams の課題の提出物など): 出力先のフォルダを選び、既定では参照先のフォルダの
 *    階層構造を出力先の下に再現して SVG を作る。「まとめる」を選ぶと、すべての SVG を出力先の直下に、
 *    階層をファイル名に含めて (課題1__田中__report.svg) 並べる
 *  - どちらも、出力先の直下に一覧 (summary.csv と index.html) を作る
 *
 * 解析の設定 (区間のしきい値・一括挿入・時間の刻み・判定ルール・有効な拡張の分類器) は、画面と同じものを使う。
 */
import {
  addTamperWarning,
  applyClassification,
  buildBuckets,
  buildDefaultTitle,
  buildFlow,
  BucketSpec,
  checkTamperEvidence,
  createAnalysisContext,
  defaultClassifiers,
  defaultRuleSet,
  describeMissingRevisions,
  extractRevisionPositions,
  extractRevisions,
  highlightCategoryMap,
  localize,
  parseDocxLayout,
  renderFlowSvg,
  renderRevisionChart,
  renderSessionedRevisionChart,
  RuleSet,
  runClassifiers,
  splitIntoSessions,
  t as libT,
  TAMPERED_SUFFIX,
} from "../../../src/core";
import type { AnalysisSettings, AppLang, BatchFolder } from "../shared";
import type { Strings } from "./strings";
import { el, extensionClassifiers, header, modal } from "./extensions";

export interface BatchContext {
  lang(): AppLang;
  strings(): Strings;
  analysis(): AnalysisSettings;
  rules(): RuleSet | undefined;
  status(msg: string): void;
}

interface Row {
  path: string;
  ok: boolean;
  chart?: string;
  flow?: string;
  revisions?: number;
  inserted?: number;
  deleted?: number;
  sessions?: number;
  highlights?: number;
  bulkInserted?: number;
  note?: string;
  /** 改ざんの痕跡が見つかったか (見つかった図は <名前>-tampered.svg にする) */
  tampered?: boolean;
}

/** 1つの docx からチャートとフローの SVG を作る */
async function renderFigures(
  bytes: Uint8Array,
  name: string,
  mtime: string,
  ctx: BatchContext,
  webUrl?: string
): Promise<Omit<Row, "path">> {
  // OneDrive / SharePoint の文書なら、見出しのファイル名を元の文書へのリンクにする
  const titleLink = webUrl ? { text: name, href: webUrl } : undefined;
  const a = ctx.analysis();
  const data = await extractRevisions(bytes);
  const missing = describeMissingRevisions(data);
  if (missing) return { ok: false, revisions: 0, note: missing };
  const model = await parseDocxLayout(bytes);
  const rules = ctx.rules() ?? defaultRuleSet(a.bulkChars);
  const actx = createAnalysisContext(model, extractRevisionPositions(model), a.gapThresholdHours);
  const extra = await extensionClassifiers(() => undefined);
  const cls = await runClassifiers([...defaultClassifiers(rules), ...extra], actx);
  applyClassification(data.events, cls);
  const bucket: BucketSpec = /^\d+$/.test(a.bucket) ? parseInt(a.bucket, 10) : (a.bucket as BucketSpec);
  const when = new Date(mtime);
  const note = ctx.rules() ? libT("rulesNote", rules.ruleSet) : undefined;
  const chartTitle = buildDefaultTitle(libT("chartTitlePrefix"), name, when);
  const chart = a.chartSplit
    ? renderSessionedRevisionChart(splitIntoSessions(data.events, a.gapThresholdHours), data.baselineCharCount, bucket, {
        title: chartTitle,
        height: 550,
        gapThresholdHours: a.gapThresholdHours,
        note,
        categories: cls.categories,
        titleLink,
      })
    : renderRevisionChart(buildBuckets(data.events, data.baselineCharCount, bucket), {
        eventRange: { start: data.events[0].date, end: data.events[data.events.length - 1].date },
        width: 1100,
        height: 550,
        title: chartTitle,
        note,
        categories: cls.categories,
        titleLink,
      });
  const flow = buildFlow(model, { gapThresholdHours: a.gapThresholdHours, bulkChars: a.bulkChars, highlightOf: highlightCategoryMap(cls) });
  const flowSvg = flow.sessions.length
    ? renderFlowSvg(flow, { title: buildDefaultTitle(libT("flowTitlePrefix"), name, when), note, categories: cls.categories, titleLink })
    : undefined;
  const tamper = await checkTamperEvidence(bytes);
  const tampered = tamper.evidence.length > 0;
  return {
    ok: true,
    chart: addTamperWarning(chart, tamper),
    flow: flowSvg && addTamperWarning(flowSvg, tamper),
    tampered,
    note: tampered ? tamper.evidence.map((e) => localize(e.message)).join(" ") : undefined,
    revisions: data.events.length,
    inserted: data.totalInserted,
    deleted: data.totalDeleted,
    sessions: flow.sessions.length,
    highlights: cls.highlights.filter((h) => cls.categories.get(h.categoryId)?.role === "highlight").length,
    bulkInserted: flow.sessions.reduce((n, s) => n + s.bulkInsChars, 0),
  };
}

/** 出力の並べ方: 階層を再現する (既定) / 出力先の直下にまとめる */
export type BatchLayout = "mirror" | "flat";

let layout: BatchLayout = "mirror";

/** "a/b/c.docx" → 出力のファイル名の元 ("a/b/c"、まとめる場合は "a__b__c") */
const stem = (p: string) => {
  const s = p.replace(/\.docx$/i, "");
  return layout === "flat" ? s.split("/").join("__") : s;
};

/** 図のファイル名 (痕跡が見つかった図は末尾に -tampered を付ける) */
const chartFile = (r: { path: string; tampered?: boolean }) => `${stem(r.path)}${r.tampered ? TAMPERED_SUFFIX : ""}.svg`;
const flowFile = (r: { path: string; tampered?: boolean }) => `${stem(r.path)}-flow${r.tampered ? TAMPERED_SUFFIX : ""}.svg`;

function csvCell(v: unknown): string {
  const s = v === undefined || v === null ? "" : String(v);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
}

/** 一覧の CSV (Excel で文字化けしないよう BOM 付き) */
function summaryCsv(rows: Row[], S: Strings): string {
  const head = S.batchColumns;
  const lines = [head.join(",")];
  for (const r of rows) {
    lines.push(
      [r.path, r.ok ? "ok" : "error", r.revisions, r.inserted, r.deleted, r.sessions, r.bulkInserted, r.highlights, r.chart ? chartFile(r) : "", r.flow ? flowFile(r) : "", r.note]
        .map(csvCell)
        .join(",")
    );
  }
  return "﻿" + lines.join("\r\n") + "\r\n";
}

/** 一覧の HTML (各 SVG へのリンク付き。この HTML 単体で見られる) */
function summaryHtml(folder: BatchFolder, rows: Row[], S: Strings, lang: AppLang): string {
  const link = (href: string, label: string) => `<a href="${escapeHtml(encodeURI(href))}">${escapeHtml(label)}</a>`;
  const body = rows
    .map(
      (r) =>
        `<tr${!r.ok || r.tampered ? ' class="err"' : ""}><td>${escapeHtml(r.path)}</td>` +
        `<td>${r.chart ? link(chartFile(r), S.batchChart) : ""}</td>` +
        `<td>${r.flow ? link(flowFile(r), S.batchFlow) : ""}</td>` +
        `<td class="n">${r.revisions ?? ""}</td><td class="n">${r.inserted ?? ""}</td><td class="n">${r.deleted ?? ""}</td>` +
        `<td class="n">${r.sessions ?? ""}</td><td class="n">${r.bulkInserted ?? ""}</td><td class="n">${r.highlights ?? ""}</td>` +
        `<td>${escapeHtml(r.note ?? "")}</td></tr>`
    )
    .join("\n");
  const cols = S.batchColumns;
  return `<!doctype html>
<html lang="${lang}"><head><meta charset="utf-8"><title>${escapeHtml(folder.name)} — ${escapeHtml(S.batchSummaryTitle)}</title>
<style>
body{font:14px/1.5 -apple-system,"Segoe UI","Hiragino Sans","Yu Gothic UI",sans-serif;margin:24px;color:#1d2026}
table{border-collapse:collapse;width:100%}th,td{border-bottom:1px solid #d9dde3;padding:6px 8px;text-align:left;vertical-align:top}
th{font-size:12px;color:#5f6672}td.n{text-align:right;font-variant-numeric:tabular-nums}tr.err td{color:#b42318}
p.note{color:#5f6672}
</style></head><body>
<h1>${escapeHtml(folder.name)}</h1>
<p>${escapeHtml(folder.location)}</p>
<p class="note">${escapeHtml(S.batchSummaryNote)}</p>
<table><thead><tr><th>${escapeHtml(cols[0])}</th><th>${escapeHtml(S.batchChart)}</th><th>${escapeHtml(S.batchFlow)}</th>${cols
    .slice(2, 8)
    .map((c) => `<th>${escapeHtml(c)}</th>`)
    .join("")}<th>${escapeHtml(cols[10])}</th></tr></thead>
<tbody>
${body}
</tbody></table>
</body></html>
`;
}

/**
 * 書き込む前の確認 (件数と、どこに何を作るか)。chooseLayout なら並べ方 (階層を再現 / まとめる) を選ばせる
 */
function confirmBatch(folder: BatchFolder, where: string, ctx: BatchContext, chooseLayout = false): Promise<boolean> {
  const S = ctx.strings();
  layout = "mirror";
  return modal<boolean>((root, close) => {
    header(root, S.batchConfirmTitle(folder.name));
    root.append(el("p", S.batchFound(folder.files.length)));
    if (folder.truncated) root.append(el("p", S.batchTruncated, "error"));
    root.append(el("p", where));
    if (chooseLayout) {
      const box = el("fieldset", undefined, "batch-layout");
      box.append(el("legend", S.batchLayout));
      for (const [value, label] of [
        ["mirror", S.batchLayoutMirror],
        ["flat", S.batchLayoutFlat],
      ] as const) {
        const l = el("label", undefined, "check-row");
        const r = el("input");
        r.type = "radio";
        r.name = "batch-layout";
        r.value = value;
        r.checked = value === "mirror";
        r.onchange = () => (layout = value);
        l.append(r, document.createTextNode(` ${label}`));
        box.append(l);
      }
      root.append(box);
    }
    root.append(el("p", S.batchOverwrite, "hint"));
    const buttons = el("div", undefined, "buttons");
    const no = el("button", S.cancel);
    no.onclick = () => close(false);
    const yes = el("button", S.batchStart, "primary");
    yes.onclick = () => close(true);
    yes.disabled = folder.files.length === 0;
    buttons.append(no, yes);
    root.append(buttons);
  }, false);
}

/**
 * 一括処理を実行する。outputRoot の下に、各ファイルのフォルダ構造を写して SVG を作る。
 * auto なら確認を省く (スモークテスト)
 */
async function runBatch(folder: BatchFolder, outputRoot: string, ctx: BatchContext): Promise<{ rows: Row[]; cancelled: boolean }> {
  const S = ctx.strings();
  const rows: Row[] = [];
  let cancelled = false;
  let closeProgress: () => void = () => undefined;
  const bar = el("progress");
  bar.max = folder.files.length;
  bar.value = 0;
  bar.className = "batch-progress";
  const current = el("p", "", "hint");
  const counter = el("p", "");
  const shown = modal<void>((root, close) => {
    closeProgress = () => close();
    header(root, S.batchRunning(folder.name));
    const buttons = el("div", undefined, "buttons");
    const stop = el("button", S.batchCancel);
    stop.onclick = () => {
      cancelled = true;
      stop.disabled = true;
    };
    buttons.append(stop);
    root.append(counter, bar, current, buttons);
  }, undefined);
  for (const [i, f] of folder.files.entries()) {
    if (cancelled) break;
    counter.textContent = S.batchProgress(i + 1, folder.files.length);
    current.textContent = f.path;
    ctx.status(`${S.batchProgress(i + 1, folder.files.length)} ${f.path}`);
    try {
      const bytes = await window.app.readBatchFile(f.ref);
      const r = await renderFigures(bytes, f.name, f.mtime, ctx, f.webUrl);
      const named = { path: f.path, tampered: r.tampered };
      if (r.chart) await window.app.writeOutput(outputRoot, chartFile(named), r.chart);
      if (r.flow) await window.app.writeOutput(outputRoot, flowFile(named), r.flow);
      rows.push({ path: f.path, ...r });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      rows.push({ path: f.path, ok: false, note: msg.replace(/^Error invoking remote method '[^']+': (Error: )?/, "") });
    }
    bar.value = i + 1;
  }
  await window.app.writeOutput(outputRoot, "summary.csv", summaryCsv(rows, S));
  await window.app.writeOutput(outputRoot, "index.html", summaryHtml(folder, rows, S, ctx.lang()));
  closeProgress();
  await shown;
  return { rows, cancelled };
}

/** 結果のまとめ (作成した数・失敗・出力先を開く) */
async function showResult(outputRoot: string, result: { rows: Row[]; cancelled: boolean }, ctx: BatchContext): Promise<void> {
  const S = ctx.strings();
  const ok = result.rows.filter((r) => r.ok).length;
  const failed = result.rows.filter((r) => !r.ok);
  ctx.status(S.batchDone(ok, failed.length, outputRoot));
  await modal<void>((root, close) => {
    header(root, result.cancelled ? S.batchCancelled : S.batchDoneTitle);
    root.append(el("p", S.batchDone(ok, failed.length, outputRoot)));
    if (failed.length) {
      const ul = el("ul", undefined, "batch-failed");
      for (const r of failed.slice(0, 20)) ul.append(el("li", `${r.path}: ${r.note ?? ""}`));
      root.append(ul);
    }
    root.append(el("p", S.batchSummaryFiles, "hint"));
    const buttons = el("div", undefined, "buttons");
    const open = el("button", S.batchOpenFolder, "primary");
    open.onclick = () => void window.app.showFolder(outputRoot);
    const done = el("button", S.close);
    done.onclick = () => close();
    buttons.append(open, done);
    root.append(buttons);
  }, undefined);
}

/** ローカルのフォルダ: 各 docx の隣に SVG を作る */
export async function batchLocalFolder(dir: string, ctx: BatchContext, auto = false): Promise<void> {
  const S = ctx.strings();
  ctx.status(S.batchListing);
  const folder = await window.app.listLocalFolder(dir);
  layout = "mirror";
  if (!auto && !(await confirmBatch(folder, S.batchWhereLocal(dir), ctx))) return;
  const result = showResult(dir, await runBatch(folder, dir, ctx), ctx);
  if (!auto) await result;
}

/**
 * SharePoint / OneDrive のフォルダ: 出力先を選び (新しいフォルダも作れる)、フォルダ構造を写して SVG を作る。
 * 失敗したら理由を返す
 */
export async function batchCloudFolder(url: string, ctx: BatchContext, auto = false): Promise<string | undefined> {
  const S = ctx.strings();
  layout = "mirror";
  ctx.status(S.batchListing);
  const r = await window.app.listCloudFolder(url);
  if (r.error || !r.folder) {
    const msg = S.urlErrors[r.error?.code ?? "network"] ?? r.error?.message ?? "";
    ctx.status(S.error(msg));
    return msg;
  }
  const folder = r.folder;
  if (folder.files.length === 0) {
    ctx.status(S.batchFound(0));
    return S.batchFound(0);
  }
  const out = await window.app.chooseOutputDir(`${folder.name}-analysis`);
  if (!out) return undefined;
  if (!auto && !(await confirmBatch(folder, S.batchWhereCloud(out), ctx, true))) return undefined;
  const result = showResult(out, await runBatch(folder, out, ctx), ctx);
  if (!auto) await result;
  return undefined;
}
