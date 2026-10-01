/**
 * レンダラ: 文書を開き、ライブラリのコア (ファイルシステムを使わない部分) で解析・描画する。
 * ファイルの読み書き・ダイアログ・設定の保存は window.app (preload) を通してメインプロセスに頼む。
 */
import {
  applyClassification,
  buildBuckets,
  buildDefaultTitle,
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
import type { AppLang, AppSettings, ExtensionListItem, OpenedFile } from "../shared";
import { STRINGS, Strings } from "./strings";
import {
  buildAnalysisResult,
  clearPanels,
  confirmPermissions,
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
  positioned: PositionedRevisionEvent[];
  classification: ClassificationResult;
  flow?: FlowResult;
  chartSvg?: string;
  flowSvg?: string;
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

  const result: Analysis = { data, positioned, classification };
  const mtime = new Date(file.mtime);
  const note = rules ? libT("rulesNote", rules.ruleSet) : undefined;
  if (data.events.length > 0) {
    const bucket: BucketSpec = /^\d+$/.test(a.bucket) ? parseInt(a.bucket, 10) : (a.bucket as BucketSpec);
    const title = buildDefaultTitle(libT("chartTitlePrefix"), file.name, mtime);
    result.chartSvg = a.chartSplit
      ? renderSessionedRevisionChart(splitIntoSessions(data.events, a.gapThresholdHours), data.baselineCharCount, bucket, {
          title,
          height: 550,
          gapThresholdHours: a.gapThresholdHours,
          note,
          categories: classification.categories,
          annotate: true,
        })
      : renderRevisionChart(buildBuckets(data.events, data.baselineCharCount, bucket), {
          eventRange: { start: data.events[0].date, end: data.events[data.events.length - 1].date },
          width: 1100,
          height: 550,
          title,
          note,
          categories: classification.categories,
          annotate: true,
        });
    const flow = buildFlow(model, {
      gapThresholdHours: a.gapThresholdHours,
      bulkChars: a.bulkChars,
      highlightOf: highlightCategoryMap(classification),
      from: parseLocal($<HTMLInputElement>("from").value),
      to: parseLocal($<HTMLInputElement>("to").value),
    });
    result.flow = flow;
    if (flow.sessions.length > 0) {
      result.flowSvg = renderFlowSvg(flow, {
        title: buildDefaultTitle(libT("flowTitlePrefix"), file.name, mtime),
        note,
        categories: classification.categories,
        annotate: true,
      });
    }
  }
  analysis = result;
  window.app.extensions.analysisComplete(
    buildAnalysisResult({
      file: { name: file.name, mtime: file.mtime },
      data,
      ctx,
      rules: ruleSet,
      classification,
      flow: result.flow,
      finalText: finalDocumentText(model),
    })
  );

  // 文書の設定・変更履歴の有無の案内
  const missing = describeMissingRevisions(data);
  if (needsHistoryFix(data.settings)) {
    banner(describeHistorySettingsProblem(data.settings), { label: S.fixHistory, run: () => void fixHistory() });
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
// 書き出し
// ---------------------------------------------------------------------------

function currentSvg(): string | undefined {
  return tab === "chart" ? analysis?.chartSvg : tab === "flow" ? analysis?.flowSvg : undefined;
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
