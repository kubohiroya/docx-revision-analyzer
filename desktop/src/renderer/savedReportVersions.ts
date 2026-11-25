import type { ReportDraft, SavedVersionReport } from "../reportData";
import { el } from "./extensions";

type Metric = "filename" | "createdAt" | "modifiedAt" | "archivedAt" | "size" | "characters" | "words" | "revisionCount";
type Column = Metric | "chart";

export interface VersionTableFigures {
  loadChart(index: number): Promise<string | undefined>;
  open(trigger: HTMLButtonElement, index: number, kind: "chart" | "flow"): void;
  signal: AbortSignal;
}

export interface VersionTableState {
  sort: Metric;
  descending: boolean;
  widths: number[];
}

const labels = {
  en: { filename: "Filename", createdAt: "Created", modifiedAt: "Modified", archivedAt: "Archived", size: "Size (B)",
    characters: "Characters", words: "Words", revisionCount: "Revisions",
    revisionCountTitle: "Tracked changes recorded in the document body, including undated changes", chart: "chart", resize: "Resize column",
    openChart: (file: string) => `Open chart for ${file}`, openFlow: (file: string) => `Open flow for ${file}`,
    compare: (file: string) => `Select ${file} for comparison`,
    selectAll: "Select all rows", clearAll: "Clear row selection", partial: "Some rows selected", none: "No rows selected",
  },
  ja: { filename: "ファイル名", createdAt: "作成日時", modifiedAt: "更新日時", archivedAt: "保管日時", size: "サイズ (B)",
    characters: "文字数", words: "単語数", revisionCount: "変更履歴",
    revisionCountTitle: "本文に記録された変更履歴の件数（日付なしを含む）", chart: "chart", resize: "列幅を調整",
    openChart: (file: string) => `${file}のチャートを表示`, openFlow: (file: string) => `${file}のフローを表示`,
    compare: (file: string) => `${file}を対象に含める`,
    selectAll: "すべての行を選択", clearAll: "すべての選択を解除", partial: "一部の行を選択", none: "選択なし",
  },
};

export function reportWordIcon(): SVGSVGElement {
  const namespace = "http://www.w3.org/2000/svg";
  const icon = document.createElementNS(namespace, "svg");
  icon.setAttribute("viewBox", "0 0 24 24");
  icon.setAttribute("aria-hidden", "true");
  icon.classList.add("report-word-icon");
  const outline = document.createElementNS(namespace, "path");
  outline.setAttribute("d", "M5 2.75h9l5 5v13.5H5zM14 2.75v5h5");
  const mark = document.createElementNS(namespace, "path");
  mark.setAttribute("d", "m7.75 11 1.7 6 2.55-5.3 2.55 5.3 1.7-6");
  icon.append(outline, mark);
  return icon;
}

export function reportVersionTable(
  drafts: readonly ReportDraft[], lang: SavedVersionReport["lang"], label: string,
  state: VersionTableState, changeSelection?: (indices: number[]) => void, figures?: VersionTableFigures,
) {
  const strings = labels[lang];
  const columns: Column[] = ["filename", "createdAt"];
  if (figures) columns.push("chart");
  columns.push("modifiedAt");
  if (drafts.some((draft) => draft.archivedAt)) columns.push("archivedAt");
  columns.push("size", "characters", "words", "revisionCount");
  if (!columns.includes(state.sort)) state.sort = "modifiedAt";
  if (state.widths.length !== columns.length) state.widths = [34, ...columns.slice(1).map(() => 66 / (columns.length - 1))];
  const wrapper = el("div", undefined, "report-draft-table-wrap");
  const table = el("table", undefined, "report-draft-table");
  table.setAttribute("aria-label", label);
  const group = document.createElement("colgroup");
  const widths = columns.map(() => el("col"));
  const selectionWidthPercent = changeSelection ? 3.3 : 0;
  if (changeSelection) {
    const selectionColumn = el("col");
    selectionColumn.style.width = `${selectionWidthPercent}%`;
    group.append(selectionColumn);
  }
  group.append(...widths);
  table.append(group);
  const header = table.createTHead().insertRow();
  const checkAll = changeSelection && el("input", undefined, "report-select-all");
  if (checkAll) {
    checkAll.type = "checkbox";
    checkAll.disabled = drafts.length === 0;
    checkAll.onchange = () => changeSelection(checkAll.checked ? drafts.map((_, index) => index) : []);
    const cell = el("th", undefined, "report-select-cell");
    cell.scope = "col";
    cell.append(checkAll);
    header.append(cell);
  }
  const body = table.createTBody();
  const collator = new Intl.Collator(lang, { numeric: true });
  const dates = new Intl.DateTimeFormat(lang, { dateStyle: "short", timeStyle: "short" });
  const numbers = new Intl.NumberFormat(lang);
  const signedNumbers = new Intl.NumberFormat(lang, { signDisplay: "exceptZero" });

  const thumbnails = new Map<HTMLElement, () => Promise<void>>();
  let selectedIndices = new Set<number>();
  function toggle(index: number): void {
    if (!changeSelection) return;
    changeSelection(selectedIndices.has(index)
      ? [...selectedIndices].filter(item => item !== index) : [...selectedIndices, index]);
  }

  function value(draft: ReportDraft, column: Metric) {
    return column === "filename" ? draft.filename ?? draft.label : draft[column];
  }

  function format(raw: string | number | undefined, column: Column): string {
    if (raw === undefined) return "";
    if (typeof raw === "number") return numbers.format(raw);
    if (column === "filename") return raw;
    const time = Date.parse(raw);
    return Number.isFinite(time) ? dates.format(time) : "";
  }

  const rows = drafts.map((draft, index) => {
    const row = el("tr");
    const name = format(value(draft, "filename"), "filename");
    const cell = el("th");
    cell.scope = "row";
    const nameLabel = el("span", name);
    nameLabel.title = name;
    const checkbox = changeSelection && el("input", undefined, "report-select-version");
    if (checkbox) {
      checkbox.type = "checkbox";
      checkbox.setAttribute("aria-label", strings.compare(name));
      checkbox.onchange = () => toggle(index);
      const selection = el("td", undefined, "report-select-cell");
      selection.append(checkbox);
      row.append(selection);
      row.classList.add("report-selectable-row");
      row.onclick = (event) => {
        if (event.target instanceof Element && !event.target.closest("button, input")) toggle(index);
      };
    }
    const file = figures ? el("button", undefined, "report-version-flow") : el("span", undefined, "report-version-name");
    file.append(reportWordIcon(), nameLabel);
    if (figures && file instanceof HTMLButtonElement) {
      file.type = "button";
      file.title = strings.openFlow(name);
      file.setAttribute("aria-label", file.title);
      file.onclick = () => figures.open(file, index, "flow");
    }
    cell.append(file);
    row.append(cell);
    for (const column of columns.slice(1)) {
      if (column === "chart") {
        if (!figures) continue;
        const target = el("td", undefined, "report-chart-cell");
        const open = el("button", undefined, "report-chart-thumbnail");
        open.type = "button";
        open.title = strings.openChart(name);
        open.setAttribute("aria-label", open.title);
        open.onclick = () => figures.open(open, index, "chart");
        open.hidden = true;
        target.append(open);
        row.append(target);
        thumbnails.set(target, async () => {
          const svg = await figures.loadChart(index);
          if (!svg || figures.signal.aborted) return;
          const image = el("img");
          image.alt = "";
          image.src = "data:image/svg+xml," + encodeURIComponent(svg);
          open.append(image);
          open.hidden = false;
        });
        continue;
      }
      const absolute = format(value(draft, column), column);
      const delta = column === "characters" ? draft.characterDelta : column === "words" ? draft.wordDelta : undefined;
      const text = absolute && delta !== undefined ? `${absolute}(${signedNumbers.format(delta)})` : absolute;
      const metric = el("td", text, column.endsWith("At") ? "report-version-date" : undefined);
      if (text) metric.title = text;
      row.append(metric);
    }
    return { row, checkbox, draft, index };
  });

  const minimum = 7;
  const headers = columns.map((column, index) => {
    const cell = el("th");
    cell.scope = "col";
    const arrow = el("span");
    arrow.setAttribute("aria-hidden", "true");
    if (column === "chart") {
      cell.classList.add("report-chart-header");
      cell.append(el("span", strings.chart));
    } else {
      const sort = el("button", undefined, "report-version-sort");
      sort.type = "button";
      if (column === "revisionCount") sort.title = strings.revisionCountTitle;
      sort.append(el("span", strings[column]), arrow);
      sort.onclick = () => {
        state.descending = state.sort === column && !state.descending;
        state.sort = column;
        order();
      };
      cell.append(sort);
    }
    header.append(cell);
    const handle = index + 1 < columns.length ? el("span", undefined, "report-column-resize") : undefined;
    if (handle) {
      handle.setAttribute("role", "separator");
      handle.setAttribute("aria-orientation", "vertical");
      handle.setAttribute("aria-label", `${strings.resize} · ${strings[column]}`);
      handle.tabIndex = 0;
      handle.onkeydown = (event) => {
        if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return;
        event.preventDefault();
        resize(index, state.widths[index] + (event.key === "ArrowLeft" ? -2 : 2));
      };
      handle.onpointerdown = (down) => {
        if (!down.isPrimary || down.button !== 0) return;
        down.preventDefault();
        handle.focus();
        handle.setPointerCapture(down.pointerId);
        const initial = state.widths[index];
        const tableWidth = table.getBoundingClientRect().width * (100 - selectionWidthPercent) / 100;
        handle.onpointermove = (move) => {
          if (handle.hasPointerCapture(move.pointerId)) resize(index, initial + 100 * (move.clientX - down.clientX) / tableWidth);
        };
        handle.onlostpointercapture = () => { handle.onpointermove = null; };
      };
      cell.append(handle);
    }
    return { cell, arrow, handle };
  });

  function resize(index: number, width: number): void {
    const total = state.widths[index] + state.widths[index + 1];
    state.widths[index] = Math.min(total - minimum, Math.max(minimum, width));
    state.widths[index + 1] = total - state.widths[index];
    applyWidths();
  }

  function applyWidths(): void {
    for (const [index, width] of widths.entries()) {
      const percent = state.widths[index] * (100 - selectionWidthPercent) / 100;
      width.style.width = `${percent}%`;
      const handle = headers[index].handle;
      handle?.setAttribute("aria-valuemin", String(minimum));
      handle?.setAttribute("aria-valuemax", String(Math.round(state.widths[index] + state.widths[index + 1] - minimum)));
      handle?.setAttribute("aria-valuenow", String(Math.round(state.widths[index])));
    }
  }

  function order(): void {
    const sorted = [...rows].sort((left, right) => {
      const a = value(left.draft, state.sort);
      const b = value(right.draft, state.sort);
      if (a === undefined || b === undefined) return Number(a === undefined) - Number(b === undefined);
      const delta = typeof a === "number" && typeof b === "number" ? a - b : collator.compare(String(a), String(b));
      return (state.descending ? -delta : delta) || left.index - right.index;
    });
    for (const [index, { cell, arrow }] of headers.entries()) {
      const active = state.sort === columns[index];
      if (active) cell.setAttribute("aria-sort", state.descending ? "descending" : "ascending");
      else cell.removeAttribute("aria-sort");
      arrow.textContent = active ? state.descending ? "▼" : "▲" : "↕";
    }
    body.replaceChildren(...sorted.map(({ row }) => row));
  }

  function update(selected: readonly number[]): void {
    selectedIndices = new Set(selected);
    for (const { index, row, checkbox } of rows) {
      const chosen = selectedIndices.has(index);
      row.dataset.selected = String(chosen);
      if (checkbox) checkbox.checked = chosen;
    }
    if (checkAll) {
      checkAll.checked = rows.length > 0 && selectedIndices.size === rows.length;
      checkAll.indeterminate = selectedIndices.size > 0 && !checkAll.checked;
      checkAll.title = checkAll.checked ? strings.selectAll : checkAll.indeterminate ? strings.partial : strings.none;
      checkAll.setAttribute("aria-label", checkAll.checked ? strings.clearAll : strings.selectAll);
    }
  }

  applyWidths();
  order();
  wrapper.append(table);
  if (figures) observeChartThumbnails(thumbnails, figures.signal);
  return { wrapper, update };
}

function observeChartThumbnails(loaders: Map<HTMLElement, () => Promise<void>>, signal: AbortSignal): void {
  const pending = new Set<HTMLElement>();
  let loading = false;
  const observer = new IntersectionObserver((entries) => {
    if (signal.aborted) return;
    for (const entry of entries) {
      const target = entry.target as HTMLElement;
      if (entry.isIntersecting && loaders.has(target)) pending.add(target);
      else pending.delete(target);
    }
    void loadVisible();
  });

  async function loadVisible(): Promise<void> {
    if (loading) return;
    loading = true;
    try {
      for (const target of pending) {
        if (signal.aborted) break;
        pending.delete(target);
        observer.unobserve(target);
        const load = loaders.get(target);
        loaders.delete(target);
        await load?.().catch(() => undefined);
      }
    } finally {
      loading = false;
    }
  }

  signal.addEventListener("abort", () => {
    observer.disconnect();
    pending.clear();
    loaders.clear();
  }, { once: true });
  for (const target of loaders.keys()) observer.observe(target);
}
