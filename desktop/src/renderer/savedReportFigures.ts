import type { AppLang, SavedLineageVersion } from "../shared";
import { renderDocumentFigures, type BatchContext } from "./batch";
import { el } from "./extensions";

export type FigureKind = "chart" | "flow";
export type FigureSettings = Pick<BatchContext, "analysis" | "rules">;

const LABELS = {
  en: {
    title: "Chart and flow", version: "Version", kind: "Figure", before: "Before", after: "After",
    chart: "Chart", flow: "Flow", back: "Back to comparison",
    loading: "Reading revision history…", unavailable: "This saved version could not be read.",
    changed: "This file has changed since it was compared. Load the versions again.",
    noFigure: "There is no figure of this type for this version.",
    failed: "The revision history could not be displayed.",
  },
  ja: {
    title: "チャート・フロー", version: "表示する版", kind: "図の種類", before: "比較前", after: "比較後",
    chart: "チャート", flow: "フロー", back: "比較に戻る",
    loading: "変更履歴を読み込んでいます…", unavailable: "この保存版を読み込めませんでした。",
    changed: "比較後にファイルが変更されました。版を読み込み直してください。",
    noFigure: "この版には選んだ種類の図がありません。",
    failed: "変更履歴を表示できませんでした。",
  },
};

export async function loadSavedVersionChart(version: SavedLineageVersion, settings: FigureSettings): Promise<string | undefined> {
  if (!version.openId) return;
  const source = await window.app.readSavedVersionFigureSource(version.openId);
  if (source.ok) return (await renderDocumentFigures(source.bytes, version.label, undefined, settings, { chartThumbnail: true })).chart;
}

export function savedReportFigures(
  lang: AppLang, versions: readonly SavedLineageVersion[], initialIndex: number, initialKind: FigureKind, settings: FigureSettings,
  close: HTMLButtonElement, back: () => void,
): HTMLElement {
  const s = LABELS[lang];
  const panel = el("section", undefined, "saved-report report-figure-panel");
  const toolbar = el("header", undefined, "report-figure-toolbar");
  const heading = el("h2", s.title);
  heading.id = "report-figure-heading";
  panel.setAttribute("aria-labelledby", heading.id);
  const backButton = el("button", s.back);
  backButton.type = "button";
  backButton.onclick = back;
  toolbar.append(backButton, heading, close);

  const controls = el("div", undefined, "report-figure-controls");
  const versionLabel = el("label", s.version);
  const select = el("select");
  const labels = versions.map((version, index) =>
    `${versions.length === 2 ? (index === 0 ? s.before + " · " : s.after + " · ") : ""}${version.label}`);
  for (const [index, label] of labels.entries()) {
    const option = el("option", label);
    option.value = String(index);
    select.append(option);
  }
  select.selectedIndex = initialIndex;
  versionLabel.append(select);
  const kinds = el("div", undefined, "report-figure-kinds");
  kinds.setAttribute("role", "group");
  kinds.setAttribute("aria-label", s.kind);
  const chart = el("button", s.chart);
  const flow = el("button", s.flow);
  chart.type = flow.type = "button";
  kinds.append(chart, flow);
  controls.append(versionLabel, kinds);

  const status = el("p", undefined, "report-figure-status");
  status.setAttribute("role", "status");
  const figure = el("div", undefined, "report-figure-viewport");
  panel.append(toolbar, controls, status, figure);

  let kind = initialKind;
  let loaded: Awaited<ReturnType<typeof renderDocumentFigures>> | undefined;
  let request = 0;
  chart.onclick = () => { kind = "chart"; drawFigure(); };
  flow.onclick = () => { kind = "flow"; drawFigure(); };
  select.onchange = () => void loadVersion();

  function drawFigure(): void {
    chart.setAttribute("aria-pressed", String(kind === "chart"));
    flow.setAttribute("aria-pressed", String(kind === "flow"));
    if (!loaded) return;
    const svg = loaded[kind];
    status.textContent = loaded.note ?? (svg ? "" : s.noFigure);
    figure.replaceChildren();
    if (!svg) return;
    const image = el("img");
    image.alt = `${s[kind]} · ${labels[select.selectedIndex]}`;
    image.src = "data:image/svg+xml," + encodeURIComponent(svg);
    figure.append(image);
  }

  async function loadVersion(): Promise<void> {
    const current = ++request;
    const version = versions[select.selectedIndex];
    loaded = undefined;
    figure.replaceChildren();
    status.textContent = s.loading;
    drawFigure();
    if (!version?.openId) {
      status.textContent = s.unavailable;
      return;
    }
    try {
      const source = await window.app.readSavedVersionFigureSource(version.openId);
      if (current !== request || !panel.isConnected) return;
      if (!source.ok) {
        status.textContent = s[source.reason];
        return;
      }
      const data = await renderDocumentFigures(source.bytes, version.label, undefined, settings);
      if (current !== request || !panel.isConnected) return;
      loaded = data;
      drawFigure();
    } catch {
      if (current === request && panel.isConnected) status.textContent = s.failed;
    }
  }

  void loadVersion();
  return panel;
}
