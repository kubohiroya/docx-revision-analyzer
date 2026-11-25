import { commentLinkUrl } from "../../../src/lib/commentReferences";
import { el } from "./extensions";
import {
  DEFAULT_REPORT_RANGE, REPORT_LIMITS, reportWindow, type ReportChange, type ReportDraft, type ReportMessage,
  type ReportPair, type ReportPassage, type ReportThread, type SavedVersionReport,
} from "../reportData";
import { reportChangeScroller } from "./reportChangeScroller";
import { reportAccordionSplitters } from "./reportAccordionSplitters";
import { reportVersionTable, reportWordIcon, type VersionTableFigures, type VersionTableState } from "./savedReportVersions";
import type { SavedVersionCommentError } from "../shared";

const EN = {
  targets: "Versions in scope",
  moreVersions: (count: number) => `${count} more`, tracking: "Versions to compare", inferred: "Inferred from saved text",
  choose: "Choose two versions to compare.", before: "Before", after: "After",
  search: "Search change text and comments", filters: "Search and filters", dates: "Filter by date", startDate: "Start date and time", endDate: "End date and time",
  clearDates: "Clear dates", invalidDates: "The end must be on or after the start.",
  dateBasis: "Dates use the saved version in which the change was observed.",
  changes: "Change list", difference: "Change details", original: "Original comment passage",
  changeCounts: (shown: string, total: string) => `${shown} matched / ${total} total`,
  characterChangeParts: (added: string, deleted: string): [string, string, string, string, string] =>
    ["Added ", added, " characters · Deleted ", deleted, " characters"],
  unknownCharacterChanges: "Character counts unavailable",
  chooseChange: "Choose a change to see its passages and comments.", export: "Export", outputRange: "Maximum text range to export",
  resizeBetween: (above: string, below: string) => `Resize ${above} and ${below}`,
  openVersion: "Open original file", author: "My Word author name", noAuthor: "Choose an author", me: "Me", latest: "Latest",
  participantCount: (n: number) => `${n} comment author${n === 1 ? "" : "s"}`, others: "Other authors",
  compose: "Add a Word comment", writeComment: "Add to Word", commentText: "Comment", commentAuthor: "Author name in Word",
  composeUnavailable: "Choose a change with one corresponding passage in a local comparison-after DOCX.",
  writingComment: "Saving the comment…",
  commentErrors: {
    disabled: "Word comment writing is disabled.", unavailable: "This Word file is no longer available. Load the versions again.",
    changed: "The Word file has changed. Load the versions again before saving.", locked: "Close the file in Word or the other application, then try again.",
    readOnly: "The Word file is read-only.", unsupported: "The selected passage or document structure does not support adding a comment.",
    saveFailed: "The Word comment could not be saved. Your draft is still here.",
  } satisfies Record<SavedVersionCommentError, string>,
  gap: "There is a gap in the saved record. The comparison uses the available endpoints.",
  missing: "A selected draft could not be read.", tracked: "Observed changes between saves, not exact edit times.",
  changeType: "Change type", tags: "Comment tag", all: "All", noMatches: "No changes match this selection.",
  evidence: "Matching evidence", candidates: (n: number) => `${n} possible sources`,
  limited: "The highlighted text shows matching edges. The middle of the passage was not fully aligned.",
  range: "Text range", ranges: ["About 20 characters", "About 80 characters", "About 200 characters", "About 400 characters", "Full page"],
  estimated: "Estimated page", page: "Page", noPage: "Page boundaries are unavailable. Showing surrounding text.",
  noPassage: "There is no corresponding passage on this side.", noText: "Text excerpts are not included.",
  paragraph: "Paragraph", sentence: "Sentence", comments: "Word comments", noComments: "This report contains no Word comments for this change.",
  unlocated: "Comments without a confirmed location", threadMissing: "This saved draft is unavailable.",
  commentAbsent: "Absent from the latest readable version", commentResolved: "Resolved in this version",
  showOriginal: "Go to the original passage", closeOriginal: "Back to the comparison",
  noOriginal: "The original range could not be located within the included text.",
  reply: "Reply", unnamed: "Unnamed author", point: "Comment position", archived: "Archived", modified: "File modified",
  moved: "Moved",
  kinds: {
    retained: "Unchanged text", rewritten: "Rewritten", split: "Split", merged: "Merged",
    copied: "Copy candidate", reappeared: "Reappearance candidate", added: "Added", removed: "Removed",
    ambiguous: "Uncertain", comment: "Comments only",
  },
  clues: {
    exact: "Matching normalized text", paragraphId: "Word paragraph ID", neighbors: "Matching neighboring text",
    heading: "Same heading", characterOverlap: "Shared character sequences", japaneseWords: "Shared Japanese word segments",
    joined: "Text matches after joining adjacent passages", past: "Identical text in an earlier draft",
  },
};

export const REPORT_LABELS: Record<SavedVersionReport["lang"], typeof EN> = { en: EN, ja: {
  targets: "対象とする版",
  moreVersions: (count) => `ほか ${count} 版`, tracking: "比較する版", inferred: "保存版の比較から推定",
  choose: "比較する版を2つ選んでください。", before: "比較前", after: "比較後",
  search: "検索：変更箇所の本文・コメント", filters: "検索とフィルタ", dates: "日時で絞り込む", startDate: "開始日時", endDate: "終了日時",
  clearDates: "日時をクリア", invalidDates: "終了日時は開始日時以降にしてください。",
  dateBasis: "変更を確認できた保存版の日時で絞り込みます。",
  changes: "変更箇所一覧", difference: "変更箇所", original: "コメントが付いた元の箇所",
  changeCounts: (shown, total) => `抽出 ${shown} 件 / 全体 ${total} 件`,
  characterChangeParts: (added, deleted) => ["追加 ", added, " 文字・削除 ", deleted, " 文字"],
  unknownCharacterChanges: "追加・削除文字数は未確定",
  chooseChange: "変更箇所を選ぶと、前後の原文とコメントを表示します。", export: "書き出す", outputRange: "書き出す原文の上限",
  resizeBetween: (above, below) => `${above}と${below}の表示高さを調整`,
  openVersion: "元のファイルを開く", author: "自分のWord作成者名", noAuthor: "作成者を選択", me: "自分", latest: "最新",
  participantCount: (n) => `${n}人がコメント`, others: "他の投稿者",
  compose: "コメントを追記", writeComment: "Wordに追記", commentText: "コメント", commentAuthor: "Wordに記録する作成者名",
  composeUnavailable: "比較後のローカルDOCXで、対応箇所が1つに特定できる変更を選んでください。",
  writingComment: "コメントを保存しています…",
  commentErrors: {
    disabled: "Wordコメントの追記は無効です。", unavailable: "このWordファイルを開けません。版を読み込み直してください。",
    changed: "Wordファイルが変更されました。版を読み込み直してから追記してください。", locked: "Wordなどでファイルを閉じてからやり直してください。",
    readOnly: "このWordファイルは読み取り専用です。", unsupported: "選択した箇所または文書の構造にはコメントを追記できません。",
    saveFailed: "Wordコメントを保存できませんでした。入力した文章は保持されています。",
  },
  gap: "この区間には保全記録の空白があります。確認できた両端を比較しています。",
  missing: "選んだ稿の一方を読み込めませんでした。", tracked: "保存版の間で観測した変化を示します。個々の編集時刻は分かりません。",
  changeType: "変更の種類", tags: "コメントのタグ", all: "すべて", noMatches: "条件に一致する変更箇所はありません。",
  evidence: "対応の手掛かり", candidates: (n) => `元の箇所の候補 ${n} 件`,
  limited: "一致する先頭と末尾を強調しています。中間部の細かな差分は特定できません。",
  range: "表示範囲", ranges: ["約20字", "約80字", "40字×5行相当", "40字×10行相当", "1ページ全体"],
  estimated: "推定ページ", page: "ページ", noPage: "ページ境界を取得できないため、前後の抜粋を表示しています。",
  noPassage: "この側に対応する箇所はありません。", noText: "原文は収録されていません。",
  paragraph: "段落", sentence: "文", comments: "Wordコメント", noComments: "この箇所のWordコメントは収録されていません。",
  unlocated: "位置を特定できないコメント", threadMissing: "元の保存版を読み込めません。",
  commentAbsent: "確認できる最新の版では消失", commentResolved: "この版で解決済み",
  showOriginal: "元の保存版の該当箇所へ", closeOriginal: "前後の比較に戻る",
  noOriginal: "収録された原文の中で元の範囲を特定できません。",
  reply: "返信", unnamed: "作成者不明", point: "コメント位置", archived: "保管日時", modified: "ファイル更新日時",
  moved: "移動",
  kinds: {
    retained: "本文は継続", rewritten: "書き換え", split: "分割", merged: "結合",
    copied: "複製候補", reappeared: "再出現候補", added: "追加", removed: "消失",
    ambiguous: "対応未確定", comment: "コメントのみ",
  },
  clues: {
    exact: "正規化した本文が一致", paragraphId: "Wordの段落ID", neighbors: "前後の文章が対応",
    heading: "同じ見出し", characterOverlap: "共通する文字列", japaneseWords: "共通する日本語の語区切り",
    joined: "隣り合う箇所をつなぐと一致", past: "以前の稿と同じ文面",
  },
} };

function button(label: string, click: () => void, cls?: string): HTMLButtonElement {
  const result = el("button", label, cls);
  result.type = "button";
  result.onclick = click;
  return result;
}

function option(label: string, value: string): HTMLOptionElement {
  const result = el("option", label);
  result.value = value;
  return result;
}

function accordion(title: string, id: string, expandedChanged?: () => void) {
  const section = el("section", undefined, "report-block");
  const header = el("div", undefined, "report-header report-accordion-header");
  const body = el("div", undefined, "report-accordion-body");
  body.id = id;
  const label = el("span", title);
  const toggle = button("", () => setExpanded(body.hidden), "report-accordion-toggle");
  const indicator = el("span", undefined, "report-accordion-chevron");
  indicator.setAttribute("aria-hidden", "true");
  toggle.id = id + "-toggle";
  toggle.setAttribute("aria-expanded", "true");
  toggle.setAttribute("aria-controls", body.id);
  toggle.append(indicator, label);
  const heading = el("h2");
  heading.append(toggle);
  header.append(heading);
  section.append(header, body);
  section.setAttribute("aria-labelledby", toggle.id);
  function setExpanded(expanded: boolean): void {
    body.hidden = !expanded;
    toggle.setAttribute("aria-expanded", String(expanded));
    expandedChanged?.();
  }
  return { section, header, body, label, toggle, setExpanded };
}

function rangeIcon(): SVGSVGElement {
  const namespace = "http://www.w3.org/2000/svg";
  const icon = document.createElementNS(namespace, "svg");
  icon.setAttribute("viewBox", "0 0 24 24");
  icon.setAttribute("aria-hidden", "true");
  const arrows = document.createElementNS(namespace, "path");
  arrows.setAttribute("d", "M3 6v12m18-12v12M7 12h10m-7-3-3 3 3 3m4-6 3 3-3 3");
  icon.append(arrows);
  return icon;
}

function reportIcon(type: "comment" | "comments" | "avatar" | "calendar" | "filter" | "search"): SVGSVGElement {
  const namespace = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(namespace, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  svg.classList.add("report-icon");
  const path = document.createElementNS(namespace, "path");
  const paths = {
    avatar: "M16 8a4 4 0 1 0-8 0 4 4 0 0 0 8 0ZM4 21a8 6 0 0 1 16 0Z",
    calendar: "M4 5h16v15H4zM8 3v4m8-4v4M4 10h16m-12 4h2m4 0h2m-8 3h2",
    comment: "M4 5h16v11H9l-5 4V5Z",
    comments: "M3 3h14v10H8l-5 4V3Zm8 13h5l5 4V8h-2",
    filter: "M4 5h16l-6 7v5l-4 2v-7Z",
    search: "M16 10a6 6 0 1 1-12 0 6 6 0 0 1 12 0Zm-1.8 4.2L20 20",
  };
  path.setAttribute("d", paths[type]);
  svg.append(path);
  return svg;
}

export function reportChangeLabel(labels: typeof EN, kind: ReportChange["kind"], moved: boolean): string {
  const join = labels === EN ? " · " : "・";
  return moved ? kind === "retained" ? labels.moved : labels.moved + join + labels.kinds[kind] : labels.kinds[kind];
}

export interface ReportViewState {
  selected?: string;
  range?: number;
  commentsOpen?: boolean;
  author?: string;
  table?: VersionTableState;
  expanded?: boolean[];
  accordionHeights?: (number | undefined)[];
  kind?: string;
  tag?: string;
  query?: string;
  filtersOpen?: boolean;
  datesOpen?: boolean;
  startDate?: string;
  endDate?: string;
}

export interface ReviewCommentDraft {
  author?: string;
  text: string;
  pending?: boolean;
  failure?: SavedVersionCommentError;
}

interface CommentComposer {
  filename: string;
  paragraph: number;
  draft: ReviewCommentDraft;
  submit(): void;
}

interface ReviewProjection {
  report: SavedVersionReport;
  sourceIndices: readonly number[];
  authors: readonly string[];
}

interface ReviewOptions {
  drafts: readonly ReportDraft[];
  targets: readonly number[];
  projection: ReviewProjection;
  project: (targets: readonly number[]) => ReviewProjection;
  included: (id: string) => boolean;
  setIncluded: (id: string, included: boolean, before: number, after: number) => void;
  outputRange: number;
  setOutputRange: (range: number) => void;
  figures: VersionTableFigures;
  openVersion: (index: number) => void;
  canOpenVersion: (index: number) => boolean;
  commentComposer?: (id: string, before: number, after: number) => CommentComposer | undefined;
}

interface ReportViewOptions {
  review?: ReviewOptions;
  state?: ReportViewState;
  initialChange?: number;
  openLink?: (url: string) => void;
}

export function renderSavedVersionReport(
  root: HTMLElement, report: SavedVersionReport, { review, state = {}, initialChange, openLink }: ReportViewOptions = {},
): void {
  const s = REPORT_LABELS[report.lang];
  state.commentsOpen ??= true;
  const drafts = review?.drafts ?? report.drafts;
  const initialTargets = report.observationTimeline ? report.drafts.map((_, index) => index) : report.initial;
  let targets = [...(review?.targets ?? initialTargets)];
  let projection = review?.projection;
  let picked = [...report.initial];
  let range = Math.min(state.range ?? DEFAULT_REPORT_RANGE, report.range);
  let selected = review ? report.pairs.flatMap(pair => pair.changes).find((change) => change.reviewId === state.selected)
    : report.observationTimeline ? report.pairs.flatMap(pair => pair.changes)[initialChange ?? 0]
    : report.pairs.find((pair) => pair.before === picked[0] && pair.after === picked[1])?.changes[initialChange ?? 0];
  let selectedPair = report.pairs.find((pair) => selected && pair.changes.includes(selected));
  let original: ReportThread | undefined;
  let originButton: HTMLButtonElement | undefined;

  const rootHeader = el("header");
  const title = el("span", report.title);
  const heading = el("h1");
  heading.append(reportWordIcon(), title);
  rootHeader.append(heading, el("p", s.tracked, "report-hint"));
  const layout = el("div", undefined, "report-layout");
  const main = el("div", undefined, "report-main");
  const sidebar = el("aside", undefined, "report-block report-comments");
  sidebar.id = "report-comments";
  sidebar.setAttribute("aria-label", s.comments);
  state.expanded ??= [true, true, true];
  const expanded = state.expanded;
  const updateExpanded = (index: number, panel: ReturnType<typeof accordion>) => {
    expanded[index] = !panel.body.hidden;
    splitters.update();
  };
  const traceTitle = review ? s.targets : s.tracking;
  const trace = accordion(traceTitle, "report-trace", () => updateExpanded(0, trace));
  const list = accordion(s.changes, "report-change-list", () => {
    updateExpanded(1, list);
    searchPanel.update();
    datePanel.update();
  });
  const detail = accordion(s.difference, "report-passages", () => {
    updateExpanded(2, detail);
    updateRangeVisibility();
  });
  const panels = [trace, list, detail];
  const splitters = reportAccordionSplitters(panels, state.accordionHeights ??= [], s.resizeBetween);
  const [versionsSplitter, passagesSplitter] = splitters.handles;
  main.append(list.section, passagesSplitter, detail.section);
  layout.append(main, sidebar);
  root.replaceChildren(rootHeader, trace.section, versionsSplitter, layout);
  root.classList.add("saved-report");

  trace.section.classList.add("report-versions");
  const selectedNames = el("span", undefined, "report-selected-names report-collapsed-summary");
  trace.toggle.append(selectedNames);
  state.table ??= { sort: "modifiedAt", descending: false, widths: [] };
  const table = reportVersionTable(drafts, report.lang, traceTitle, state.table,
    review || drafts.length > 2 ? changeTargets : undefined, review?.figures);
  trace.body.append(table.wrapper);

  function changeTargets(next: number[]): void {
    targets = [...next].sort((a, b) => a - b);
    if (review) {
      projection = review.project(targets);
      report = projection.report;
      picked = [...report.initial];
    } else {
      picked = targets.length >= 2 ? [targets[0], targets[targets.length - 1]] : targets;
    }
    selected = undefined;
    selectedPair = undefined;
    original = undefined;
    draw(true);
  }

  const formatCount = new Intl.NumberFormat(report.lang);
  const formatSignedCount = new Intl.NumberFormat(report.lang, { signDisplay: "always" });
  const listCounts = el("span", undefined, "report-hint report-heading-counts report-list-counts");
  listCounts.setAttribute("aria-live", "polite");
  list.label.after(listCounts);
  const selectedTitle = el("span", undefined, "report-selected-title report-collapsed-summary");
  list.toggle.append(selectedTitle);
  const filters = el("div", undefined, "report-filters");
  filters.id = "report-change-filters";
  const kinds = el("select");
  const kindLabel = el("label", s.changeType);
  kindLabel.append(kinds);
  const tags = el("select");
  const tagLabel = el("label", s.tags);
  tagLabel.append(tags);
  const status = el("p", undefined, "report-hint");
  status.setAttribute("role", "status");
  const results = el("ul", undefined, "report-changes");
  const changeScroll = el("div", undefined, "report-change-scroll");
  changeScroll.tabIndex = 0;
  changeScroll.setAttribute("role", "region");
  changeScroll.setAttribute("aria-label", s.changes);
  changeScroll.append(results);
  const search = el("input");
  search.type = "search";
  search.value = state.query ?? "";
  search.placeholder = s.search;
  search.setAttribute("aria-label", s.search);
  const searchLabel = el("label", undefined, "report-change-search");
  searchLabel.append(reportIcon("search"), search);
  filters.append(searchLabel, kindLabel, tagLabel);

  const period = el("div", undefined, "report-date-period");
  period.id = "report-date-period";
  const fromDate = el("input");
  const toDate = el("input");
  fromDate.type = toDate.type = "datetime-local";
  fromDate.step = toDate.step = "1";
  fromDate.value = state.startDate ?? "";
  toDate.value = state.endDate ?? "";
  const fromLabel = el("label", s.startDate);
  const toLabel = el("label", s.endDate);
  fromLabel.append(fromDate);
  toLabel.append(toDate);
  const resetDates = button(s.clearDates, () => {
    fromDate.value = toDate.value = "";
    applyDates();
  });
  const dateBasis = el("span", s.dateBasis);
  dateBasis.id = "report-date-basis";
  dateBasis.hidden = true;
  for (const field of [fromDate, toDate]) {
    field.title = s.dateBasis;
    field.setAttribute("aria-describedby", dateBasis.id);
  }
  period.append(fromLabel, toLabel, resetDates, dateBasis);

  const searchPanel = changePanel(filters, search, "filtersOpen", "filter", s.filters, "report-filter-toggle");
  const datePanel = changePanel(period, fromDate, "datesOpen", "calendar", s.dates, "report-date-toggle");
  const changeTools = el("div", undefined, "report-change-tools");
  changeTools.append(searchPanel.toggle, datePanel.toggle);
  list.header.append(changeTools);
  list.body.append(filters, period, status, changeScroll);
  const changeRows = reportChangeScroller<ChangeEntry>(changeScroll, results, renderChangeRow);
  const searchIndex = new WeakMap<ReportChange, { tags: Set<string>; text: string }>();

  function changePanel(
    panel: HTMLElement, field: HTMLInputElement, key: "filtersOpen" | "datesOpen",
    icon: "filter" | "calendar", label: string, className: string,
  ) {
    const toggle = button("", () => {
      state[key] = list.body.hidden || panel.hidden;
      if (list.body.hidden) list.setExpanded(true);
      update();
      if (state[key]) field.focus({ preventScroll: true });
    }, className);
    toggle.append(reportIcon(icon));
    toggle.setAttribute("aria-label", label);
    toggle.setAttribute("aria-controls", panel.id);
    toggle.title = label;
    panel.onkeydown = (event) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      event.stopPropagation();
      state[key] = false;
      update();
      toggle.focus({ preventScroll: true });
    };

    function update(): void {
      panel.hidden = !state[key] || toggle.disabled;
      toggle.setAttribute("aria-expanded", String(!panel.hidden && !list.body.hidden));
    }

    return { toggle, update };
  }

  function applyDates(): void {
    state.startDate = fromDate.value;
    state.endDate = toDate.value;
    draw(true);
  }
  fromDate.oninput = toDate.oninput = applyDates;
  search.oninput = () => {
    state.query = search.value;
    draw(true);
  };

  const rangeMenu = el("details", undefined, "report-range-menu");
  const rangeToggle = el("summary", undefined, "report-range-toggle");
  rangeToggle.setAttribute("aria-label", s.range);
  rangeToggle.append(rangeIcon(), el("span", s.range));
  const context = el("div", undefined, "report-range");
  const rangeInput = el("input");
  rangeInput.type = "range";
  rangeInput.min = "0";
  rangeInput.max = String(report.range);
  rangeInput.step = "1";
  rangeInput.value = String(range);
  const rangeValue = el("output");
  const rangeName = el("span", s.range);
  const rangeControl = el("label");
  rangeControl.append(rangeName, rangeValue, rangeInput);
  const endpoints = el("div", undefined, "report-header report-hint");
  endpoints.append(el("span", s.ranges[0]), el("span", report.range ? s.ranges[report.range] : ""));
  context.append(rangeControl, endpoints);
  if (review) {
    const outputLabel = el("label", s.outputRange, "report-output-range");
    const output = el("select");
    output.append(...s.ranges.map((description, index) => option(description, String(index))));
    output.value = String(review.outputRange);
    output.onchange = () => review.setOutputRange(Number(output.value));
    outputLabel.append(output);
    context.append(outputLabel);
  }
  rangeMenu.append(rangeToggle, context);
  const back = button(s.closeOriginal, () => {
    original = undefined;
    drawPassages();
    originButton?.focus({ preventScroll: true });
  });
  const characterCounts = el("span", undefined, "report-hint report-heading-counts report-character-counts");
  detail.label.after(characterCounts);
  const comparisonKind = el("span", undefined, "report-hint report-accordion-meta");
  detail.toggle.append(comparisonKind);
  const commentToggle = button("", () => {
    state.commentsOpen = !state.commentsOpen;
    drawThreads();
  }, "report-comment-toggle");
  commentToggle.setAttribute("aria-label", s.comments);
  commentToggle.setAttribute("aria-controls", sidebar.id);
  commentToggle.title = s.comments;
  commentToggle.append(reportIcon("comment"));
  const tools = el("div", undefined, "report-passage-tools");
  tools.append(rangeMenu, commentToggle);
  const detailToolbar = el("div", undefined, "report-detail-toolbar");
  detailToolbar.append(back, tools);
  const passages = el("div", undefined, "report-passage-content");
  detail.body.append(detailToolbar, passages);

  rangeMenu.ontoggle = () => {
    if (rangeMenu.open && !rangeInput.disabled) rangeInput.focus({ preventScroll: true });
  };
  rangeMenu.onkeydown = (event) => {
    if (event.key !== "Escape" || !rangeMenu.open) return;
    event.preventDefault();
    event.stopPropagation();
    rangeMenu.open = false;
    rangeToggle.focus({ preventScroll: true });
  };
  rangeMenu.addEventListener("focusout", (event) => {
    if (event.relatedTarget instanceof Node && !rangeMenu.contains(event.relatedTarget)) rangeMenu.open = false;
  });
  const closeRangeOnOutsidePress = (event: PointerEvent) => {
    if (event.target instanceof Node && !rangeMenu.contains(event.target)) rangeMenu.open = false;
  };
  rootHeader.onpointerdown = closeRangeOnOutsidePress;
  layout.onpointerdown = closeRangeOnOutsidePress;

  const changeFilters = () => { state.kind = kinds.value; state.tag = tags.value; original = undefined; draw(true); };
  kinds.onchange = changeFilters;
  tags.onchange = changeFilters;
  rangeInput.oninput = () => {
    state.range = range = Number(rangeInput.value);
    drawPassages();
  };

  const sorted = () => [...picked].sort((a, b) => a - b);
  const pairFor = () => {
    const [before, after] = sorted();
    return selectedPair ?? (report.observationTimeline ? report.pairs[0]
      : report.pairs.find((pair) => pair.before === before && pair.after === after));
  };
  const rangeLabel = (p: ReportPassage) => [p.heading, `${s.paragraph} ${p.paragraph}`].filter(Boolean).join(" · ");
  const threadLabel = (thread: ReportThread) => thread.draft === undefined ? thread.sourceLabel ?? "" : report.drafts[thread.draft].label;
  const kindValue = (change: ReportChange) => change.reordered ? `moved:${change.kind}` : change.kind;
  const changeName = (change: ReportChange) => reportChangeLabel(s, change.kind, change.reordered);

  function displayDate(date: string): string {
    const time = new Date(date);
    return Number.isFinite(time.valueOf()) ? time.toLocaleString(report.lang) : "";
  }

  function changeTitle(change: ReportChange): string {
    const left = change.before[0];
    const right = change.after[0];
    return left && right && left.heading !== right.heading ? `${rangeLabel(left)} → ${rangeLabel(right)}` : rangeLabel(right ?? left);
  }

  function updateSelectedTitle(): void {
    const title = selected ? changeTitle(selected) : "";
    selectedTitle.textContent = selectedTitle.title = title;
  }

  function avatar(own: boolean): HTMLElement {
    const badge = el("span", undefined, own ? "report-author-avatar" : "report-author-avatar report-other-author");
    badge.append(reportIcon("avatar"));
    return badge;
  }

  function participants(change: ReportChange): HTMLElement | undefined {
    const authors = new Set(change.threads.flatMap((id) => report.threads[id].messages.map((message) => message.author?.trim() || "")));
    if (!authors.size) return;
    const own = Boolean(state.author && authors.has(state.author));
    const others = authors.size > Number(own);
    const roles = review ? [own && s.me, others && s.others].filter(Boolean) : [];
    const label = [s.participantCount(authors.size), ...roles].join(" · ");
    const node = el("span", undefined, "report-participants");
    node.setAttribute("role", "img");
    node.setAttribute("aria-label", label);
    node.title = label;
    if (review && own) node.append(avatar(true));
    node.append(reportIcon(authors.size > 1 ? "comments" : "comment"));
    if (review && others) node.append(avatar(false));
    return node;
  }

  type ChangeEntry = { item: ReportChange; pair: ReportPair; observed?: string };

  function searchText(text: string): string {
    return text.normalize("NFKC").toLocaleLowerCase(report.lang);
  }

  function indexChange(item: ReportChange) {
    const cached = searchIndex.get(item);
    if (cached) return cached;
    const messages = item.threads.flatMap(id => report.threads[id].messages);
    const tags = new Set(messages.flatMap(message => message.tags));
    let normalized: string | undefined;
    const indexed = {
      tags,
      get text(): string {
        if (normalized !== undefined) return normalized;
        const content = item.searchText ?? [...item.before, ...item.after].flatMap(passage => {
          const excerpt = passage.excerpt;
          if (!excerpt) return [];
          const text = report.clips[excerpt.clip].text;
          return [excerpt.passage ? text.slice(...excerpt.passage) : text];
        }).join("\n");
        normalized = searchText([changeTitle(item), content, ...messages.map(message => message.text), ...tags].join("\n"));
        return normalized;
      },
    };
    searchIndex.set(item, indexed);
    return indexed;
  }

  function draw(reset = false): void {
    rangeMenu.open = false;
    if (report.title) title.textContent = report.title;
    table.update(targets);
    const visibleNames = targets.slice(0, 2).map((index) => {
      const name = drafts[index].filename ?? drafts[index].label;
      const chip = el("span", name);
      chip.title = name;
      return chip;
    });
    if (targets.length > 2) visibleNames.push(el("span", s.moreVersions(targets.length - 2)));
    selectedNames.replaceChildren(...visibleNames);
    const [before, after] = sorted();
    const pairs = report.observationTimeline
      ? report.pairs.filter(pair => review || (targets.includes(pair.before) && targets.includes(pair.after)))
      : report.pairs.filter(pair => pair.before === (review ? report.initial[0] : before) && pair.after === (review ? report.initial[1] : after));
    const entries: ChangeEntry[] = pairs.flatMap(pair => {
      const version = report.drafts[pair.after];
      const observed = version.archivedAt ?? version.modifiedAt ?? version.createdAt;
      return pair.changes.map(item => ({ item, pair, observed }));
    });
    let kind = state.kind ?? "";
    let tag = state.tag ?? "";
    const availableKinds = new Map(entries.map(({ item }) => [kindValue(item), changeName(item)]));
    const availableTags = new Set(entries.flatMap(({ item }) => [...indexChange(item).tags]));
    if (!availableKinds.has(kind)) kind = "";
    if (!availableTags.has(tag)) tag = "";
    kinds.replaceChildren(option(s.all, ""), ...[...availableKinds].map(([value, label]) => option(label, value)));
    tags.replaceChildren(option(s.all, ""), ...[...availableTags].map((value) => option(`#${value}`, value)));
    kinds.value = state.kind = kind;
    tags.value = state.tag = tag;
    const terms = searchText(state.query ?? "").trim().split(/\s+/u).filter(Boolean);
    const from = state.startDate ? Date.parse(state.startDate) : -Infinity;
    const to = state.endDate ? Date.parse(state.endDate) : Infinity;
    const dateActive = Boolean(state.startDate || state.endDate);
    searchPanel.toggle.dataset.active = String(Boolean(terms.length || kind || tag));
    datePanel.toggle.dataset.active = String(dateActive);
    const invalid = from > to;
    fromDate.setAttribute("aria-invalid", String(invalid));
    toDate.setAttribute("aria-invalid", String(invalid));
    const shown = invalid ? [] : entries.filter(({ item, observed }) => {
      const date = Date.parse(observed ?? "");
      const indexed = indexChange(item);
      return (!kind || kindValue(item) === kind) && (!tag || indexed.tags.has(tag)) &&
        (!dateActive || (date >= from && date <= to)) && terms.every(term => indexed.text.includes(term));
    });
    if (selected && !shown.some(({ item }) => item === selected)) {
      selected = undefined;
      selectedPair = undefined;
      original = undefined;
    }
    state.selected = selected?.reviewId;
    searchPanel.toggle.disabled = datePanel.toggle.disabled = !pairs.length;
    searchPanel.update();
    datePanel.update();
    changeScroll.hidden = !shown.length;
    const gap = pairs.some(pair => pair.gap);
    status.hidden = !invalid && Boolean(shown.length && !gap);
    status.textContent = invalid ? s.invalidDates : !pairs.length ? s.choose
      : gap ? s.gap : pairs.every(pair => pair.unavailable) ? s.missing : s.noMatches;
    listCounts.textContent = s.changeCounts(formatCount.format(shown.length), formatCount.format(entries.length));
    updateSelectedTitle();
    changeRows.setItems(shown, reset);
    drawPassages();
    drawThreads();
    splitters.update();
  }

  function renderChangeRow({ item, pair, observed }: ChangeEntry): HTMLLIElement {
    const title = changeTitle(item);
    const row = button("", () => {
      rangeMenu.open = false;
      selected = selected === item ? undefined : item;
      selectedPair = selected ? pair : undefined;
      state.selected = selected?.reviewId;
      original = undefined;
      for (const previous of results.querySelectorAll("li > button")) {
        previous.setAttribute("aria-pressed", String(previous === row && selected === item));
      }
      updateSelectedTitle();
      drawPassages();
      drawThreads();
    });
    row.setAttribute("aria-pressed", String(item === selected));
    const heading = el("span", undefined, "report-change-heading");
    const label = el("span", title, "report-change-title");
    label.title = title;
    heading.append(label);
    if (item.characterChanges) {
      const added = formatSignedCount.format(item.characterChanges.added);
      const deleted = formatSignedCount.format(-item.characterChanges.deleted);
      const counts = el("span", undefined, "report-row-character-counts");
      counts.title = s.characterChangeParts(added, deleted).join("");
      counts.append(
        el("span", added, "report-added-count"),
        el("span", deleted, "report-deleted-count"),
      );
      heading.append(counts);
    }
    const authors = participants(item);
    if (authors) heading.append(authors);
    const when = observed ? displayDate(observed) : "";
    row.append(heading, el("small", [changeName(item), when].filter(Boolean).join(" · ")));
    const first = item.after[0] ?? item.before[0];
    const excerpt = first.excerpt;
    if (excerpt) {
      const text = report.clips[excerpt.clip].text;
      const [start, end] = reportWindow(text, excerpt.focus, 44);
      row.append(el("span", (start ? "…" : "") + text.slice(start, end).replace(/\s+/gu, " ") + (end < text.length ? "…" : ""), "report-row-text"));
    }
    const li = el("li");
    li.append(row);
    const id = item.reviewId;
    if (review && id && projection) {
      const before = projection.sourceIndices[pair.before];
      const after = projection.sourceIndices[pair.after];
      const label = el("label", undefined, "report-include-change");
      const check = el("input");
      check.type = "checkbox";
      check.checked = review.included(id);
      check.setAttribute("aria-label", `${s.export} · ${title}`);
      check.onchange = () => review.setIncluded(id, check.checked, before, after);
      label.append(check, s.export);
      li.append(label);
    }
    return li;
  }

  function passageView(passage: ReportPassage, mode: "before" | "after" | "original" | "match"): HTMLElement {
    const node = el("article", undefined, "report-passage");
    node.append(el("p", rangeLabel(passage), "report-hint"));
    const excerpt = passage.excerpt;
    if (!excerpt) {
      node.append(el("p", s.noText, "report-hint"));
      return node;
    }
    const clip = report.clips[excerpt.clip];
    const [start, end] = reportWindow(clip.text, excerpt.focus, mode === "original" ? Infinity : REPORT_LIMITS[range]);
    if (REPORT_LIMITS[range] === Infinity) {
      const page = clip.page;
      node.append(el("p", page ? `${page.estimated ? s.estimated : s.page} ${page.number}` : s.noPage, "report-hint"));
    }
    const text = el("p", undefined, "report-text");
    if (start || clip.beforeCut) text.append("…");
    let offset = start;
    const tag = mode === "before" ? "del" : mode === "after" ? "ins" : "mark";
    for (const [left, right] of excerpt.marks) {
      if (excerpt.point && start <= left && left <= end) {
        const caret = el("mark", "│");
        caret.setAttribute("aria-label", s.point);
        text.append(clip.text.slice(offset, left), caret);
        offset = left;
      } else if (right > start && left < end) {
        const a = Math.max(offset, left, start);
        const b = Math.min(end, right);
        if (a < b) {
          text.append(clip.text.slice(offset, a), el(tag, clip.text.slice(a, b)));
          offset = b;
        }
      }
    }
    text.append(clip.text.slice(offset, end));
    if (end < clip.text.length || clip.afterCut) text.append("…");
    node.append(text);
    return node;
  }

  function drawPassages(): void {
    detail.label.textContent = original ? s.original : s.difference;
    const changes = original ? undefined : selected?.characterChanges;
    if (changes) {
      const [prefix, added, join, deleted, suffix] = s.characterChangeParts(
        formatSignedCount.format(changes.added), formatSignedCount.format(-changes.deleted),
      );
      characterCounts.replaceChildren(
        prefix, el("span", added, "report-added-count"), join, el("span", deleted, "report-deleted-count"), suffix,
      );
    } else {
      characterCounts.textContent = changes === null ? s.unknownCharacterChanges : "";
    }
    comparisonKind.textContent = selected ? changeName(selected) : "";
    updateRangeVisibility();
    const detailBody = passages;
    detailBody.replaceChildren();
    if (original) {
      detailBody.append(el("h3", threadLabel(original)));
      for (const passage of original.original ?? []) detailBody.append(passageView(passage, "original"));
      addOpenVersion(detailBody, original.reviewVersion);
      return;
    }
    const pair = pairFor();
    if (!selected || !pair) {
      detailBody.append(el("p", s.chooseChange, "report-hint"));
      return;
    }
    const change = selected;
    const { before, after } = pair;
    rangeValue.textContent = s.ranges[range];
    rangeToggle.title = `${s.range} · ${s.ranges[range]}`;
    rangeInput.setAttribute("aria-valuetext", s.ranges[range]);
    rangeInput.disabled = !report.range;
    const evidence = change.evidence.map((clue) => s.clues[clue]).join(" / ");
    if (evidence) detailBody.append(el("p", `${s.evidence} · ${evidence}`, "report-hint"));
    if (change.alignmentPartial) detailBody.append(el("p", s.limited, "report-hint"));
    if (change.candidateCount) detailBody.append(el("p", s.candidates(change.candidateCount), "report-hint"));
    const columns = el("div", undefined, "report-columns");
    for (const [side, version, passages] of [["before", before, change.before], ["after", after, change.after]] as const) {
      const column = el("section");
      column.append(el("h3", `${s[side]} · ${report.drafts[version].label}`));
      if (!passages.length) column.append(el("p", s.noPassage, "report-hint"));
      for (const passage of passages) column.append(passageView(passage, change.alignmentPartial ? "match" : side));
      if (passages.length) addOpenVersion(column, projection?.sourceIndices[version]);
      columns.append(column);
    }
    detailBody.append(columns);
  }

  function addOpenVersion(container: HTMLElement, index?: number): void {
    if (index !== undefined && review?.canOpenVersion(index)) {
      container.append(button(s.openVersion, () => review.openVersion(index), "report-open-version"));
    }
  }

  function updateRangeVisibility(): void {
    rangeMenu.hidden = Boolean(original) || !report.includeText || !selected || detail.body.hidden;
    if (rangeMenu.hidden) rangeMenu.open = false;
    back.hidden = !original || detail.body.hidden;
  }

  function showThread(thread: ReportThread, allowTagFilter: boolean, latest?: ReportMessage): HTMLElement {
    const node = el("article", undefined, "report-thread");
    node.append(el("h3", threadLabel(thread)));
    const notes = [thread.absent && s.commentAbsent, thread.resolved && s.commentResolved, thread.unavailable && s.threadMissing].filter(Boolean);
    if (notes.length) node.append(el("p", notes.join(" · "), "report-hint"));
    for (const [index, message] of thread.messages.entries()) {
      const post = el("div", undefined, "report-post");
      if (index) post.append(el("p", s.reply, "report-hint"));
      const own = Boolean(review && state.author && state.author === message.author?.trim());
      const author = el("strong", undefined, "report-post-author");
      if (review) author.append(avatar(own));
      author.append(message.author || s.unnamed);
      if (own) author.append(el("small", s.me));
      if (own && message === latest) {
        post.classList.add("report-latest-own");
        author.append(el("small", s.latest, "report-latest-marker"));
      }
      post.append(author);
      if (message.date) post.append(el("p", displayDate(message.date), "report-hint"));
      post.append(el("p", message.text, "report-text"));
      const chips = el("div", undefined, "report-tags");
      for (const tag of message.tags) {
        if (allowTagFilter) {
          chips.append(button(`#${tag}`, () => { tags.value = tag; changeFilters(); tags.focus(); }));
        } else {
          chips.append(el("span", `#${tag}`));
        }
      }
      if (message.tags.length) post.append(chips);
      for (const reference of message.links ?? []) {
        const url = commentLinkUrl(reference.url);
        if (!url) continue;
        const link = el("a", reference.label || url);
        link.href = url;
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        if (openLink) link.onclick = (event) => { event.preventDefault(); openLink(url); };
        const row = el("p");
        row.append(link);
        post.append(row);
      }
      node.append(post);
    }
    if (thread.original?.length) {
      const show = button(s.showOriginal, () => {
        original = thread;
        originButton = show;
        detail.setExpanded(true);
        drawPassages();
        back.focus({ preventScroll: true });
        detail.section.scrollIntoView({ block: "nearest" });
      }, "report-original");
      node.append(show);
    } else {
      node.append(el("p", s.noOriginal, "report-hint"));
    }
    return node;
  }

  function drawThreads(): void {
    const reportVersions = selectedPair ? [selectedPair.before, selectedPair.after] : review ? report.initial : sorted();
    const unlocated = reportVersions.flatMap((id) => report.drafts[id]?.unlocated ?? []);
    sidebar.hidden = state.commentsOpen === false;
    layout.classList.toggle("report-comments-open", !sidebar.hidden);
    commentToggle.setAttribute("aria-expanded", String(!sidebar.hidden));
    if (sidebar.hidden) return;
    sidebar.replaceChildren(el("h2", s.comments));
    if (projection?.authors.length) {
      const label = el("label", s.author, "report-my-author");
      const selector = el("select");
      selector.append(option(s.noAuthor, ""), ...projection.authors.map((author) => option(author, author)));
      selector.value = state.author ?? "";
      selector.onchange = () => {
        state.author = selector.value;
        draw();
        sidebar.querySelector("select")?.focus({ preventScroll: true });
      };
      label.append(selector);
      sidebar.append(label);
    }
    if (selected?.threads.length) {
      const threads = selected.threads.map((id) => report.threads[id]);
      let latest = threads[0].messages.at(-1);
      for (const thread of threads.slice(1)) {
        const message = thread.messages.at(-1);
        if (message?.date && latest?.date && Date.parse(message.date) > Date.parse(latest.date)) latest = message;
      }
      for (const thread of threads) sidebar.append(showThread(thread, true, latest));
    } else if (selected) {
      sidebar.append(el("p", s.noComments, "report-hint"));
    } else {
      sidebar.append(el("p", s.chooseChange, "report-hint"));
    }
    if (unlocated.length) {
      const box = el("details");
      box.append(el("summary", s.unlocated));
      for (const id of unlocated) box.append(showThread(report.threads[id], false));
      sidebar.append(box);
    }
    if (selected?.reviewId && selectedPair && review?.commentComposer && projection) {
      const { sourceIndices } = projection;
      const compose = review.commentComposer(selected.reviewId, sourceIndices[selectedPair.before], sourceIndices[selectedPair.after]);
      sidebar.append(compose ? commentForm(compose) : el("p", s.composeUnavailable, "report-hint"));
    }
  }

  function commentForm({ filename, paragraph, draft, submit }: CommentComposer): HTMLElement {
    const form = el("section", undefined, "report-comment-compose");
    form.append(el("h3", s.compose), el("p", `${filename} · ${s.paragraph} ${paragraph}`, "report-hint"));
    const authorLabel = el("label", s.commentAuthor);
    const author = el("input");
    author.type = "text";
    author.autocomplete = "off";
    author.value = draft.author ?? state.author ?? "";
    authorLabel.append(author);
    const textLabel = el("label", s.commentText);
    const text = el("textarea");
    text.rows = 3;
    text.value = draft.text;
    textLabel.append(text);
    const save = button(s.writeComment, () => {
      draft.author = author.value;
      draft.text = text.value;
      submit();
      update();
    }, "report-write-comment");
    const status = el("p", undefined, "report-hint");
    status.setAttribute("role", "status");
    form.append(authorLabel, textLabel, save, status);
    author.oninput = text.oninput = () => {
      draft.author = author.value;
      draft.text = text.value;
      update();
    };
    function update(): void {
      author.disabled = text.disabled = draft.pending === true;
      save.disabled = Boolean(draft.pending || !author.value.trim() || !text.value.trim());
      status.textContent = draft.pending ? s.writingComment : draft.failure ? s.commentErrors[draft.failure] : "";
    }
    update();
    return form;
  }

  draw();
  for (const [index, section] of panels.entries()) section.setExpanded(expanded[index]);
  requestAnimationFrame(splitters.update);
}
