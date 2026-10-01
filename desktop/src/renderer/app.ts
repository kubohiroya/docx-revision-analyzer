/**
 * レンダラ: 文書を開き、ライブラリのコア (ファイルシステムを使わない部分) で解析・描画する。
 * ファイルの読み書き・ダイアログ・設定の保存は window.app (preload) を通してメインプロセスに頼む。
 */
import {
  applyClassification,
  buildBuckets,
  buildDefaultTitle,
  checkIntegrity,
  chartTargets,
  FigureTarget,
  flowTargets,
  ResolvedAnnotation,
  sessionedChartTargets,
  IntegrityReport,
  buildFlow,
  BucketSpec,
  ClassificationResult,
  createAnalysisContext,
  defaultClassifiers,
  defaultRuleSet,
  describeHistorySettingsProblem,
  describeMissingRevisions,
  DocxRevisionData,
  extractRevisionPositions,
  extractRevisions,
  finalDocumentText,
  FlowResult,
  highlightCategoryMap,
  Highlight,
  needsHistoryFix,
  parseDocxLayout,
  parseRuleSet,
  PositionedRevisionEvent,
  renderFlowSvg,
  renderRevisionChart,
  renderSessionedRevisionChart,
  RuleSet,
  RuleSetError,
  runClassifiers,
  setLang,
  splitIntoSessions,
  t as libT,
} from "../../../src/core";
import type { AppLang, AppSettings, ExtensionListItem, MicrosoftStatus, OpenedFile, RecentUrl } from "../shared";
import { STRINGS, Strings } from "./strings";
import {
  annotationsForExport,
  bindFigureAnnotations,
  collectFigureAnnotations,
  FigureAnnotations,
} from "./figures";
import {
  buildAnalysisResult,
  clearPanels,
  confirmPermissions,
  el,
  header,
  modal,
  extensionClassifiers,
  initExtensionUi,
  loc,
  permissionList,
  renderPanels,
  showSendLog,
} from "./extensions";

type Tab = "chart" | "flow" | "highlights" | "settings";

interface Analysis {
  data: DocxRevisionData;
  integrity: IntegrityReport;
  positioned: PositionedRevisionEvent[];
  classification: ClassificationResult;
  flow?: FlowResult;
  chartSvg?: string;
  flowSvg?: string;
  /** 注釈を埋め込んで描き直す (書き出し用) */
  renderChart?: (annotations?: Map<string, ResolvedAnnotation>) => string;
  renderFlow?: (annotations?: Map<string, ResolvedAnnotation>) => string;
  chartTargets: FigureTarget[];
  flowTargets: FigureTarget[];
  /** 拡張の図の注釈 (図 → 部分のキー → 注釈) */
  figureAnnotations: FigureAnnotations;
}

interface Internal {
  ready(): void;
  smokeAnalyzed(summary: { ok: boolean; message: string }): void;
}

const appInternal = (window as unknown as { appInternal: Internal }).appInternal;

let settings: AppSettings;
let lang: AppLang = "en";
let S: Strings = STRINGS.en;
let file: OpenedFile | undefined;
let rules: RuleSet | undefined;
let rulesError: string | undefined;
let analysis: Analysis | undefined;
let tab: Tab = "chart";
let smokeReported = false;
let updatesAvailable = false;
let extList: ExtensionListItem[] = [];
/** 解析中に起きた拡張のエラー (案内に表示する) */
let extErrors: string[] = [];

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

function setText(id: string, text: string): void {
  $(id).textContent = text;
}

// ---------------------------------------------------------------------------
// 表示言語
// ---------------------------------------------------------------------------

function applyLang(): void {
  S = STRINGS[lang];
  setLang(lang);
  document.documentElement.lang = lang;
  setText("open", S.open);
  setText("open-url", S.openUrl);
  setText("open-url2", S.openUrl);
  setText("tab-chart", S.tabChart);
  setText("tab-flow", S.tabFlow);
  setText("tab-highlights", S.tabHighlights);
  setText("tab-settings", S.tabSettings);
  setText("l-gap", S.gap);
  setText("l-from", S.from);
  setText("l-to", S.to);
  setText("l-bucket", S.bucket);
  setText("o-auto", S.bucketAuto);
  setText("o-minute", S.bucketMinute);
  setText("o-hour", S.bucketHour);
  setText("o-day", S.bucketDay);
  setText("l-split", S.chartSplit);
  setText("save-svg", S.saveSvg);
  setText("save-png", S.savePng);
  setText("drop-here", S.dropHere);
  setText("open2", S.open);
  setText("drop-hint", S.dropHint);
  setText("purpose", S.purpose);
  setText("s-lang", S.settingsLang);
  setText("o-lang-system", S.langSystem);
  setText("s-rules", S.settingsRules);
  setText("rules-choose", S.chooseRules);
  setText("rules-clear", S.clearRules);
  setText("s-ext", S.settingsExtensions);
  setText("s-updates", S.settingsUpdates);
  setText("l-updates", S.checkForUpdates);
  setText("s-about", S.settingsAbout);
  setText("about-text", S.aboutText);
  if (!file) setText("file-name", S.noFile);
  renderSettings();
}

// ---------------------------------------------------------------------------
// タブ
// ---------------------------------------------------------------------------

function showTab(next: Tab): void {
  tab = next;
  if (next === "settings") {
    void window.app.extensions.list().then((l) => {
      extList = l;
      renderExtensions();
    });
  }
  for (const b of document.querySelectorAll<HTMLButtonElement>(".tabs button")) {
    b.setAttribute("aria-selected", String(b.dataset.tab === next));
  }
  const hasFile = !!analysis;
  $("welcome").hidden = hasFile || next === "settings";
  for (const t of ["chart", "flow", "highlights", "settings"] as Tab[]) {
    $(`pane-${t}`).hidden = t !== next || (t !== "settings" && !hasFile);
  }
  $("controls").hidden = !hasFile || next === "settings";
  for (const el of document.querySelectorAll<HTMLElement>(".chart-only")) el.hidden = next !== "chart";
  for (const el of document.querySelectorAll<HTMLElement>(".flow-only")) el.hidden = next !== "flow";
  for (const el of document.querySelectorAll<HTMLElement>(".figure-only")) el.hidden = next !== "chart" && next !== "flow";
}

(window as unknown as { __showTab: (t: Tab) => void }).__showTab = showTab;

// ---------------------------------------------------------------------------
// 解析と描画
// ---------------------------------------------------------------------------

function status(msg: string): void {
  setText("status", msg);
}

function banner(text: string | undefined, action?: { label: string; run: () => void }): void {
  $("banner").hidden = !text;
  setText("banner-text", text ?? "");
  const btn = $<HTMLButtonElement>("banner-action");
  btn.hidden = !action;
  if (action) {
    btn.textContent = action.label;
    btn.onclick = action.run;
  }
}

function parseLocal(v: string): Date | undefined {
  if (!v) return undefined;
  const d = new Date(v);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

function readControls(): void {
  const a = settings.analysis;
  const gap = parseFloat($<HTMLInputElement>("gap").value);
  if (Number.isFinite(gap) && gap > 0) a.gapThresholdHours = gap;
  const bulk = parseFloat($<HTMLInputElement>("bulk").value);
  if (Number.isFinite(bulk) && bulk > 0) a.bulkChars = bulk;
  a.bucket = $<HTMLSelectElement>("bucket").value;
  a.chartSplit = $<HTMLInputElement>("split").checked;
}

function writeControls(): void {
  const a = settings.analysis;
  $<HTMLInputElement>("gap").value = String(a.gapThresholdHours);
  $<HTMLInputElement>("bulk").value = String(a.bulkChars);
  $<HTMLInputElement>("bulk").disabled = !!rules;
  $<HTMLInputElement>("bulk").title = rules ? S.bulkCharsByRules : "";
  setText("l-bulk", rules ? `${S.bulkChars} — ${S.bulkCharsByRules}` : S.bulkChars);
  $<HTMLSelectElement>("bucket").value = a.bucket;
  $<HTMLInputElement>("split").checked = a.chartSplit;
}

async function analyze(): Promise<void> {
  if (!file) return;
  status(S.analyzing);
  const a = settings.analysis;
  const bytes = file.bytes;
  const data = await extractRevisions(bytes);
  const model = await parseDocxLayout(bytes);
  const positioned = extractRevisionPositions(model);
  const ruleSet = rules ?? defaultRuleSet(a.bulkChars);
  const ctx = createAnalysisContext(model, positioned, a.gapThresholdHours);
  extErrors = [];
  clearPanels();
  const classification = await runPipeline(ruleSet, ctx);
  applyClassification(data.events, classification);

  const result: Analysis = {
    data,
    positioned,
    classification,
    integrity: await checkIntegrity(bytes),
    chartTargets: [],
    flowTargets: [],
    figureAnnotations: { chart: new Map(), flow: new Map() },
  };
  const mtime = new Date(file.mtime);
  const fileName = file.name;
  const note = rules ? libT("rulesNote", rules.ruleSet) : undefined;
  if (data.events.length > 0) {
    const bucket: BucketSpec = /^\d+$/.test(a.bucket) ? parseInt(a.bucket, 10) : (a.bucket as BucketSpec);
    const title = buildDefaultTitle(libT("chartTitlePrefix"), file.name, mtime);
    const highlightIds = classification.categories.highlights().map((c) => c.id);
    if (a.chartSplit) {
      const sessions = splitIntoSessions(data.events, a.gapThresholdHours);
      result.renderChart = (annotations) =>
        renderSessionedRevisionChart(sessions, data.baselineCharCount, bucket, {
          title,
          height: 550,
          gapThresholdHours: a.gapThresholdHours,
          note,
          categories: classification.categories,
          annotate: !annotations,
          annotations,
        });
      result.chartTargets = sessionedChartTargets(sessions, data.baselineCharCount, bucket, highlightIds);
    } else {
      const buckets = buildBuckets(data.events, data.baselineCharCount, bucket);
      result.renderChart = (annotations) =>
        renderRevisionChart(buckets, {
          eventRange: { start: data.events[0].date, end: data.events[data.events.length - 1].date },
          width: 1100,
          height: 550,
          title,
          note,
          categories: classification.categories,
          annotate: !annotations,
          annotations,
        });
      result.chartTargets = chartTargets(buckets, highlightIds);
    }
    result.chartSvg = result.renderChart();
    const flow = buildFlow(model, {
      gapThresholdHours: a.gapThresholdHours,
      bulkChars: a.bulkChars,
      highlightOf: highlightCategoryMap(classification),
      from: parseLocal($<HTMLInputElement>("from").value),
      to: parseLocal($<HTMLInputElement>("to").value),
    });
    result.flow = flow;
    if (flow.sessions.length > 0) {
      result.renderFlow = (annotations) =>
        renderFlowSvg(flow, {
          title: buildDefaultTitle(libT("flowTitlePrefix"), fileName, mtime),
          note,
          categories: classification.categories,
          annotate: !annotations,
          annotations,
        });
      result.flowSvg = result.renderFlow();
      result.flowTargets = flowTargets(flow);
    }
  }
  analysis = result;
  const analysisResult = buildAnalysisResult({
      file: { name: file.name, mtime: file.mtime },
      data,
      ctx,
      rules: ruleSet,
      classification,
      flow: result.flow,
      finalText: finalDocumentText(model),
      integrity: result.integrity,
    });
  window.app.extensions.analysisComplete(analysisResult);
  result.figureAnnotations = await collectFigureAnnotations(
    { chart: result.chartTargets, flow: result.flowTargets },
    analysisResult,
    (msg) => extErrors.push(msg)
  );

  // 文書の設定・変更履歴の有無の案内
  const missing = describeMissingRevisions(data);
  if (needsHistoryFix(data.settings)) {
    // OneDrive / SharePoint の文書はアプリから書き換えられないため、案内だけを出す
    if (file.source?.kind === "url") banner(`${describeHistorySettingsProblem(data.settings)} ${S.cloudFixHint}`);
    else banner(describeHistorySettingsProblem(data.settings), { label: S.fixHistory, run: () => void fixHistory() });
  } else if (missing) {
    banner(missing);
  } else if (rulesError) {
    banner(S.rulesError(rulesError));
  } else if (extErrors.length) {
    banner(extErrors.join(" / "));
  } else {
    banner(undefined);
  }

  $("pane-chart").innerHTML = result.chartSvg ?? "";
  $("pane-flow").innerHTML = result.flowSvg ?? (data.events.length > 0 ? `<p class="hint">${escapeHtml(libT("noRevisionsInRange"))}</p>` : "");
  bindFigureAnnotations($("pane-chart"), result.figureAnnotations.chart);
  bindFigureAnnotations($("pane-flow"), result.figureAnnotations.flow);
  renderHighlights();
  showTab(tab === "settings" ? "chart" : tab);
  status(
    `${file.name} — ${data.events.length} revisions, ${classification.highlights.length} highlights` +
      (classification.classifiers.length ? ` (${classification.classifiers.map((c) => `${c.id}@${c.version}`).join(", ")})` : "")
  );
  if (!smokeReported) {
    smokeReported = true;
    appInternal.smokeAnalyzed({ ok: true, message: `${data.events.length} events, ${classification.highlights.length} highlights` });
  }
}

/** 既定の分類器 (並べ替え・判定ルール) のパイプライン。拡張の分類器 (#15) はここに加える */
async function runPipeline(ruleSet: RuleSet, ctx: ReturnType<typeof createAnalysisContext>): Promise<ClassificationResult> {
  const fromExtensions = await extensionClassifiers((msg) => extErrors.push(msg));
  return runClassifiers([...defaultClassifiers(ruleSet), ...fromExtensions], ctx);
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
}

function localized(v: Highlight["reason"]): string {
  if (typeof v === "string") return v;
  return v[lang] ?? v.en ?? v.ja ?? "";
}

function fmtTime(d: Date): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

/** ハイライトの挿入の本文 (最終文書での順。挿入の間は " … " でつなぐ)。すべて後で削除されていれば印を付ける */
function excerptOf(h: Highlight, a: Analysis): string {
  const max = 160;
  const ids = new Set(h.eventIds);
  const events = a.positioned
    .filter((e) => e.type === "ins" && ids.has(e.id))
    .sort((x, y) => x.final.start.docOffset - y.final.start.docOffset);
  const text = events.map((e) => e.text.trim()).filter(Boolean).join(" … ");
  const clipped = text.length > max ? `${text.slice(0, max)}…` : text;
  return events.every((e) => e.finalChars === 0) ? `${S.excerptDeleted} ${clipped}` : clipped;
}

function renderHighlights(): void {
  const pane = $("pane-highlights");
  const a = analysis;
  if (!a) {
    pane.innerHTML = "";
    return;
  }
  const items = [...a.classification.highlights].sort((x, y) => x.timeRange.start.getTime() - y.timeRange.start.getTime());
  if (items.length === 0) {
    pane.innerHTML = `<p class="hint">${escapeHtml(S.highlightsEmpty)}</p>`;
    renderPanels(pane);
    renderIntegrity(pane, a.integrity);
    return;
  }
  const reg = a.classification.categories;
  const rows = items.map((h, i) => {
    const cat = reg.get(h.categoryId);
    const color = cat ? cat.color : "#888888";
    const label = cat ? reg.label(cat) : h.categoryId;
    return `<tr class="item" data-index="${i}">
      <td>${escapeHtml(fmtTime(h.timeRange.start))}</td>
      <td><span class="swatch" style="background:${color}"></span>${escapeHtml(label)}</td>
      <td>${escapeHtml(localized(h.reason))}</td>
      <td class="excerpt">${escapeHtml(excerptOf(h, a))}</td>
      <td><div class="goto"><button data-goto="chart">${escapeHtml(S.showInChart)}</button><button data-goto="flow">${escapeHtml(S.showInFlow)}</button></div></td>
    </tr>`;
  });
  pane.innerHTML = `<p class="hint">${escapeHtml(S.highlightsIntro)}</p>
    <table class="highlights"><thead><tr><th>${escapeHtml(S.colWhen)}</th><th>${escapeHtml(S.colCategory)}</th><th>${escapeHtml(S.colReason)}</th><th>${escapeHtml(S.colExcerpt)}</th><th></th></tr></thead>
    <tbody>${rows.join("")}</tbody></table>`;
  for (const tr of pane.querySelectorAll<HTMLTableRowElement>("tr.item")) {
    const h = items[Number(tr.dataset.index)];
    tr.addEventListener("click", (e) => {
      const target = (e.target as HTMLElement).closest<HTMLButtonElement>("button[data-goto]");
      for (const r of pane.querySelectorAll("tr.selected")) r.classList.remove("selected");
      tr.classList.add("selected");
      if (target?.dataset.goto === "chart") locateInChart(h);
      else locateInFlow(h);
    });
  }
  renderPanels(pane);
  renderIntegrity(pane, a.integrity);
}

/** 整合性の簡易チェック (判定ではなく情報として表示する) */
function renderIntegrity(pane: HTMLElement, report: IntegrityReport): void {
  const box = document.createElement("section");
  box.className = "ext-panel integrity";
  const h = document.createElement("h3");
  h.textContent = S.integrityTitle;
  const note = document.createElement("p");
  note.className = "hint";
  note.textContent = localized(report.note);
  const ul = document.createElement("ul");
  for (const item of report.items) {
    const li = document.createElement("li");
    const tag = document.createElement("span");
    tag.className = "hint";
    tag.textContent = ` (${item.observed === null ? S.integrityNotAvailable : item.observed ? S.integrityObserved : S.integrityOk})`;
    li.append(document.createTextNode(localized(item.message)), tag);
    ul.append(li);
  }
  box.append(h, note, ul);
  pane.append(box);
}

function clearFocus(): void {
  for (const el of document.querySelectorAll(".dra-focus")) el.classList.remove("dra-focus");
}

/** チャートで、ハイライトの時刻を含む棒を示す */
function locateInChart(h: Highlight): void {
  showTab("chart");
  clearFocus();
  const t = h.timeRange.start.getTime();
  const bars = [...$("pane-chart").querySelectorAll<SVGElement>("[data-start]")].filter((el) => {
    const s = Date.parse(el.dataset.start!);
    const e = Date.parse(el.dataset.end!);
    return s <= t && t < e;
  });
  for (const b of bars) b.classList.add("dra-focus");
  bars[0]?.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
}

/** フローで、ハイライトの区間の終了時点の列の、対象の段落を示す */
function locateInFlow(h: Highlight): void {
  showTab("flow");
  clearFocus();
  const a = analysis;
  if (!a?.flow) return;
  const t = h.timeRange.start.getTime();
  const k = a.flow.sessions.findIndex((s) => s.start.getTime() <= t && t <= s.end.getTime());
  if (k < 0) return;
  const ids = new Set(h.eventIds);
  const paras = new Set(a.positioned.filter((e) => e.type === "ins" && ids.has(e.id)).flatMap((e) => e.paraModelIndices));
  const column = $("pane-flow").querySelector(`[data-column="${k + 1}"]`);
  const els = [...(column?.querySelectorAll<SVGElement>("[data-para]") ?? [])].filter((el) => paras.has(Number(el.dataset.para)));
  for (const el of els) el.classList.add("dra-focus");
  els[0]?.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
}

// ---------------------------------------------------------------------------
// ファイル
// ---------------------------------------------------------------------------

async function openFile(opened: OpenedFile | null): Promise<void> {
  if (!opened) return;
  file = opened;
  setText("file-name", opened.path);
  document.title = `${opened.name} — ${S.appTitle}`;
  await runSafely(analyze);
}

async function fixHistory(): Promise<void> {
  if (!file) return;
  await runSafely(async () => {
    const r = await window.app.preserveHistory(file!.path);
    file = await window.app.readDocx(file!.path);
    await analyze();
    banner(S.fixed(r.backupPath ?? ""));
  });
}

async function runSafely(fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    status(S.error(msg));
    banner(S.error(msg));
    if (!smokeReported) {
      smokeReported = true;
      appInternal.smokeAnalyzed({ ok: false, message: msg });
    }
  }
}

// ---------------------------------------------------------------------------
// OneDrive / SharePoint の URL から開く
// ---------------------------------------------------------------------------

let msStatus: MicrosoftStatus = { configured: false };
let recentUrls: RecentUrl[] = [];

function hostOf(url: string): string {
  try {
    return new URL(url).hostname;
  } catch {
    return url;
  }
}

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleString(lang === "ja" ? "ja-JP" : "en-US", { dateStyle: "medium", timeStyle: "short" });
}

/** 最近開いた文書の一覧 (クリックで開く、× で一覧から削除) */
function recentList(onOpen: (r: RecentUrl) => void, limit = recentUrls.length): HTMLElement {
  const box = el("div");
  if (recentUrls.length === 0) {
    box.append(el("p", S.urlRecentEmpty, "hint"));
    return box;
  }
  for (const r of recentUrls.slice(0, limit)) {
    const row = el("div", undefined, "recent-item");
    const open = el("button", undefined, "open");
    open.title = r.url;
    open.append(el("span", r.name), el("small", `${hostOf(r.url)} — ${fmtDate(r.openedAt)}`));
    open.onclick = () => onOpen(r);
    const remove = el("button", "×", "remove");
    remove.title = S.urlRemove;
    remove.setAttribute("aria-label", `${S.urlRemove}: ${r.name}`);
    remove.onclick = async (e) => {
      e.stopPropagation();
      recentUrls = await window.app.removeRecentUrl(r.url);
      row.remove();
      renderRecentWelcome();
    };
    row.append(open, remove);
    box.append(row);
  }
  return box;
}

/** 起動画面の「最近開いた文書」 */
function renderRecentWelcome(): void {
  const box = $("recent-welcome");
  box.replaceChildren();
  if (!msStatus.configured || recentUrls.length === 0) return;
  box.append(el("h3", S.urlRecent), recentList((r) => void openFromUrl({ recent: r.url }), 5));
}

async function refreshCloud(): Promise<void> {
  [msStatus, recentUrls] = await Promise.all([window.app.microsoftStatus(), window.app.recentUrls()]);
  renderRecentWelcome();
  renderMicrosoftSettings();
}

/** URL を開く。失敗したら理由を返す (ダイアログに表示するため) */
async function openFromUrl(target: { url: string } | { recent: string }): Promise<string | undefined> {
  status(S.urlOpening);
  const r = await window.app.openUrl(target);
  await refreshCloud();
  if (r.error) {
    const msg = S.urlErrors[r.error.code] ?? r.error.message;
    status(S.error(msg));
    return r.error.message && r.error.code !== "badUrl" ? `${msg} (${r.error.message})` : msg;
  }
  await openFile(r.file ?? null);
  return undefined;
}

/** 「URL から開く」ダイアログ: URL の入力欄と最近開いた文書の一覧 */
async function showUrlDialog(): Promise<void> {
  await refreshCloud();
  await modal<void>((root, close) => {
    header(root, S.urlDialogTitle);
    if (!msStatus.configured) {
      root.append(el("p", S.urlNotConfigured, "error"));
    } else {
      root.append(el("p", S.urlDialogHelp, "hint"));
      const input = el("input");
      input.type = "url";
      input.className = "url-input mono";
      input.placeholder = S.urlPlaceholder;
      input.spellcheck = false;
      const error = el("p", "", "error");
      error.hidden = true;
      const go = el("button", S.urlOpen, "primary");
      const run = async (target: { url: string } | { recent: string }) => {
        go.disabled = true;
        error.hidden = true;
        const msg = await openFromUrl(target);
        go.disabled = false;
        if (msg) {
          error.textContent = msg;
          error.hidden = false;
        } else close();
      };
      go.onclick = () => void (input.value.trim() && run({ url: input.value.trim() }));
      input.onkeydown = (e) => {
        if (e.key === "Enter") go.click();
      };
      const row = el("div", undefined, "row");
      row.append(input, go);
      root.append(row, error);
      const account = el("p", msStatus.account ? S.urlSignedInAs(msStatus.account) : S.urlNotSignedIn, "hint");
      if (msStatus.account) {
        const out = el("button", S.signOut, "link");
        out.onclick = async () => {
          msStatus = await window.app.microsoftSignOut();
          account.textContent = S.urlNotSignedIn;
          out.remove();
          renderMicrosoftSettings();
        };
        account.append(" ", out);
      }
      root.append(account, el("h3", S.urlRecent), recentList((r) => void run({ recent: r.url })));
      setTimeout(() => input.focus(), 0);
    }
    const buttons = el("div", undefined, "buttons");
    const cancel = el("button", S.close);
    cancel.onclick = () => close();
    buttons.append(cancel);
    root.append(buttons);
  }, undefined);
}

function renderMicrosoftSettings(): void {
  if (!settings) return;
  setText("s-ms", S.settingsMicrosoft);
  setText(
    "ms-status",
    !msStatus.configured ? S.msNotConfigured : msStatus.account ? S.urlSignedInAs(msStatus.account) : S.urlNotSignedIn
  );
  const out = $<HTMLButtonElement>("ms-signout");
  out.textContent = S.signOut;
  out.hidden = !msStatus.account;
  setText("ms-advanced", S.msAdvanced);
  setText("ms-client-help", S.msClientHelp);
  setText("ms-client-save", S.save);
  $<HTMLInputElement>("ms-client-id").value = settings.microsoft?.clientId ?? "";
  $<HTMLInputElement>("ms-client-id").placeholder = "00000000-0000-0000-0000-000000000000";
}

// ---------------------------------------------------------------------------
// 書き出し
// ---------------------------------------------------------------------------

/** 書き出す SVG: 拡張の注釈を <title> (説明) と <a href> (リンク) として埋め込む */
function currentSvg(): string | undefined {
  const a = analysis;
  if (!a) return undefined;
  if (tab === "chart" && a.renderChart) return a.renderChart(annotationsForExport(a.figureAnnotations.chart, lang));
  if (tab === "flow" && a.renderFlow) return a.renderFlow(annotationsForExport(a.figureAnnotations.flow, lang));
  return undefined;
}

function svgToPng(svg: string, scale = 2): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(img.width * scale);
      canvas.height = Math.ceil(img.height * scale);
      const g = canvas.getContext("2d")!;
      g.fillStyle = "#ffffff";
      g.fillRect(0, 0, canvas.width, canvas.height);
      g.scale(scale, scale);
      g.drawImage(img, 0, 0);
      URL.revokeObjectURL(url);
      canvas.toBlob(async (blob) => {
        if (!blob) return reject(new Error("PNG conversion failed"));
        resolve(new Uint8Array(await blob.arrayBuffer()));
      }, "image/png");
    };
    img.onerror = () => reject(new Error("SVG could not be loaded"));
    img.src = url;
  });
}

async function save(kind: "svg" | "png"): Promise<void> {
  const svg = currentSvg();
  if (!svg || !file) return;
  const base = `${file.name.replace(/\.docx$/i, "")}-${tab}.${kind}`;
  await runSafely(async () => {
    const out = await window.app.saveFile(base, kind === "svg" ? svg : await svgToPng(svg), kind);
    if (out) status(S.saved(out));
  });
}

// ---------------------------------------------------------------------------
// 設定
// ---------------------------------------------------------------------------

async function persist(): Promise<void> {
  await window.app.setSettings(settings);
}

async function loadRules(): Promise<void> {
  rules = undefined;
  rulesError = undefined;
  if (!settings.rulesPath) return;
  try {
    rules = parseRuleSet((await window.app.readRules(settings.rulesPath)).value);
  } catch (err) {
    rulesError = err instanceof RuleSetError || err instanceof Error ? err.message : String(err);
  }
}

function renderSettings(): void {
  if (!settings) return;
  $<HTMLSelectElement>("lang").value = settings.lang ?? "";
  setText("rules-current", settings.rulesPath ?? S.rulesDefault(settings.analysis.bulkChars));
  $("rules-error").hidden = !rulesError;
  setText("rules-error", rulesError ? S.rulesError(rulesError) : "");
  $<HTMLButtonElement>("rules-clear").disabled = !settings.rulesPath;
  $<HTMLInputElement>("updates").checked = !!settings.checkForUpdates;
  $<HTMLInputElement>("updates").disabled = !updatesAvailable;
  setText("updates-note", updatesAvailable ? S.updatesNote : `${S.updatesNote} ${S.updatesDevBuild}`);
  renderExtensions();
  renderMicrosoftSettings();
  renderRecentWelcome();
}

function renderExtensions(): void {
  const box = $("ext-list");
  box.replaceChildren();
  if (extList.length === 0) {
    box.innerHTML = `<p class="hint">${escapeHtml(S.extensionsNone)}</p>`;
    return;
  }
  const intro = document.createElement("p");
  intro.className = "hint";
  intro.textContent = S.extensionsIntro;
  box.append(intro);
  for (const item of extList) {
    const div = document.createElement("div");
    div.className = "ext-item";
    const head = document.createElement("div");
    head.className = "head";
    const title = document.createElement("strong");
    title.textContent = `${loc(item.manifest?.name ?? item.id, lang)} ${item.manifest ? `v${item.manifest.version}` : ""}`;
    const toggle = document.createElement("label");
    toggle.className = "check-row";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = item.enabled;
    cb.disabled = !item.manifest;
    cb.onchange = () =>
      void runSafely(async () => {
        if (cb.checked && !(await confirmPermissions(item))) {
          cb.checked = false;
          return;
        }
        extList = await window.app.extensions.setEnabled(item.id, cb.checked);
        renderExtensions();
        if (file) await analyze();
      });
    toggle.append(cb, document.createTextNode(` ${S.extEnable}`));
    head.append(title, toggle);
    div.append(head);
    const desc = document.createElement("p");
    desc.className = "hint";
    desc.textContent = [
      loc(item.manifest?.description, lang),
      item.state === "active" ? S.extStateActive : item.state === "starting" ? S.extStateStarting : "",
      item.error ? S.extStateError(item.error) : "",
    ]
      .filter(Boolean)
      .join(" — ");
    div.append(desc);
    const permTitle = document.createElement("span");
    permTitle.className = "hint";
    permTitle.textContent = S.extPermissions;
    div.append(permTitle, permissionList(item));
    if (item.manifest?.permissions?.network?.length) {
      const row = document.createElement("div");
      row.className = "row";
      const confirm = document.createElement("label");
      confirm.className = "check-row";
      const ccb = document.createElement("input");
      ccb.type = "checkbox";
      ccb.checked = settings.extensionConfirmSends?.[item.id] !== false;
      ccb.onchange = () => {
        settings.extensionConfirmSends = { ...(settings.extensionConfirmSends ?? {}), [item.id]: ccb.checked };
        void persist();
      };
      confirm.append(ccb, document.createTextNode(` ${S.extConfirmSends}`));
      const logBtn = document.createElement("button");
      logBtn.textContent = S.extSendLog;
      logBtn.onclick = () => void showSendLog(item);
      row.append(confirm, logBtn);
      div.append(row);
    }
    box.append(div);
  }
}

// ---------------------------------------------------------------------------
// 起動
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  settings = await window.app.getSettings();
  initExtensionUi({ lang: () => lang, strings: () => S, onPanelsChanged: () => renderHighlights() });
  extList = await window.app.extensions.list();
  updatesAvailable = await window.app.updatesAvailable();
  lang = settings.lang ?? (await window.app.systemLang());
  await loadRules();
  applyLang();
  writeControls();
  setText("about-version", S.version(await window.app.version()));

  $("open").onclick = () => void runSafely(async () => openFile(await window.app.openDocxDialog()));
  $("open2").onclick = () => $("open").click();
  $("open-url").onclick = () => void showUrlDialog();
  $("open-url2").onclick = () => void showUrlDialog();
  $("ms-signout").onclick = () =>
    void runSafely(async () => {
      msStatus = await window.app.microsoftSignOut();
      renderMicrosoftSettings();
    });
  $("ms-client-save").onclick = () =>
    void runSafely(async () => {
      const id = $<HTMLInputElement>("ms-client-id").value.trim();
      settings.microsoft = { ...(settings.microsoft ?? {}), clientId: id || undefined };
      await persist();
      await refreshCloud();
    });
  window.app.onOpenUrl((url) => void openFromUrl({ url }));
  (window as unknown as { __showUrlDialog: () => void }).__showUrlDialog = () => void showUrlDialog();
  await refreshCloud();
  for (const b of document.querySelectorAll<HTMLButtonElement>(".tabs button")) {
    b.onclick = () => showTab(b.dataset.tab as Tab);
  }
  let timer: number | undefined;
  for (const id of ["gap", "from", "to", "bulk", "bucket", "split"]) {
    $(id).addEventListener("change", () => {
      readControls();
      void persist();
      window.clearTimeout(timer);
      timer = window.setTimeout(() => void runSafely(analyze), 150);
    });
  }
  $("save-svg").onclick = () => void save("svg");
  $("save-png").onclick = () => void save("png");

  $("lang").addEventListener("change", async () => {
    const v = $<HTMLSelectElement>("lang").value as AppLang | "";
    settings.lang = v || undefined;
    lang = settings.lang ?? (await window.app.systemLang());
    await persist();
    applyLang();
    writeControls();
    setText("about-version", S.version(await window.app.version()));
    if (file) await runSafely(analyze);
  });
  $("updates").addEventListener("change", () => {
    settings.checkForUpdates = $<HTMLInputElement>("updates").checked;
    void persist();
  });
  $("rules-choose").onclick = () =>
    void runSafely(async () => {
      const r = await window.app.chooseRulesDialog();
      if (!r) return;
      settings.rulesPath = r.path;
      await persist();
      await loadRules();
      renderSettings();
      writeControls();
      if (file) await analyze();
    });
  $("rules-clear").onclick = () =>
    void runSafely(async () => {
      settings.rulesPath = undefined;
      await persist();
      await loadRules();
      renderSettings();
      writeControls();
      if (file) await analyze();
    });

  // ドラッグ&ドロップ
  document.addEventListener("dragover", (e) => {
    e.preventDefault();
    document.body.classList.add("dragover");
  });
  document.addEventListener("dragleave", () => document.body.classList.remove("dragover"));
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    document.body.classList.remove("dragover");
    const f = e.dataTransfer?.files[0];
    if (f) void runSafely(async () => openFile(await window.app.readDocx(window.app.pathForFile(f))));
  });
  window.app.onOpenPath((p) => void runSafely(async () => openFile(await window.app.readDocx(p))));

  showTab("chart");
  appInternal.ready();
}

void main();
