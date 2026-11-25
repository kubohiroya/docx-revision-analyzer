import type { ReportDraft, SavedVersionReport } from "../reportData";
import { el } from "./extensions";

type Column = "filename" | "createdAt" | "modifiedAt" | "archivedAt" | "size" | "characters" | "words";

export interface VersionTableState {
  sort: Column;
  descending: boolean;
  widths: number[];
}

const labels = {
  en: { filename: "Filename", createdAt: "Created", modifiedAt: "Modified", archivedAt: "Archived", size: "Size (B)",
    characters: "Characters", words: "Words", resize: "Resize column" },
  ja: { filename: "ファイル名", createdAt: "作成日時", modifiedAt: "更新日時", archivedAt: "保管日時", size: "サイズ (B)",
    characters: "文字数", words: "単語数", resize: "列幅を調整" },
};

export function reportVersionTable(
  drafts: readonly ReportDraft[], lang: SavedVersionReport["lang"], label: string,
  state: VersionTableState, select?: (index: number) => void,
) {
  const strings = labels[lang];
  const columns: Column[] = ["filename", "createdAt", "modifiedAt"];
  if (drafts.some((draft) => draft.archivedAt)) columns.push("archivedAt");
  columns.push("size", "characters", "words");
  if (!columns.includes(state.sort)) state.sort = "modifiedAt";
  if (state.widths.length !== columns.length) state.widths = [34, ...columns.slice(1).map(() => 66 / (columns.length - 1))];
  const wrapper = el("div", undefined, "report-draft-table-wrap");
  const table = el("table", undefined, "report-draft-table");
  table.setAttribute("aria-label", label);
  const group = document.createElement("colgroup");
  const widths = columns.map(() => el("col"));
  group.append(...widths);
  table.append(group);
  const header = table.createTHead().insertRow();
  const body = table.createTBody();
  const collator = new Intl.Collator(lang, { numeric: true });
  const dates = new Intl.DateTimeFormat(lang, { dateStyle: "short", timeStyle: "short" });
  const numbers = new Intl.NumberFormat(lang);

  function value(draft: ReportDraft, column: Column) {
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
    const role = el("small");
    const control = select ? el("button", undefined, "report-version-pick") : el("span", undefined, "report-version-pick");
    if (control instanceof HTMLButtonElement) {
      control.type = "button";
      control.onclick = () => select?.(index);
      control.setAttribute("aria-pressed", "false");
    }
    control.append(nameLabel, role);
    cell.append(control);
    row.append(cell);
    for (const column of columns.slice(1)) {
      const text = format(value(draft, column), column);
      const metric = el("td", text);
      if (text) metric.title = text;
      row.append(metric);
    }
    return { row, control, role, draft, index };
  });

  const minimum = 7;
  const headers = columns.map((column, index) => {
    const cell = el("th");
    cell.scope = "col";
    const sort = el("button", undefined, "report-version-sort");
    sort.type = "button";
    const arrow = el("span");
    arrow.setAttribute("aria-hidden", "true");
    sort.append(el("span", strings[column]), arrow);
    sort.onclick = () => {
      state.descending = state.sort === column && !state.descending;
      state.sort = column;
      order();
    };
    cell.append(sort);
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
        const tableWidth = table.getBoundingClientRect().width;
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
      width.style.width = `${state.widths[index]}%`;
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

  function update(selected: readonly number[], before: string, after: string, one: string): void {
    for (const { index, row, control, role } of rows) {
      const chosen = selected.includes(index);
      row.dataset.selected = String(chosen);
      if (select) control.setAttribute("aria-pressed", String(chosen));
      role.textContent = chosen ? selected.length === 1 ? one : index === selected[0] ? before : after : "";
    }
  }

  applyWidths();
  order();
  wrapper.append(table);
  return { wrapper, update };
}
