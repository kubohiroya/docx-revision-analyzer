import { commentLinkUrl } from "../../../src/lib/commentReferences";
import { el } from "./extensions";
import {
  DEFAULT_REPORT_RANGE, REPORT_LIMITS, reportWindow, type ReportChange, type ReportDraft, type ReportMessage,
  type ReportPassage, type ReportThread, type SavedVersionReport,
} from "../reportData";
import { reportVersionTable, type VersionTableState } from "./savedReportVersions";
import type { SavedVersionCommentError } from "../shared";

const EN = {
  tracking: "Versions to compare", inferred: "Inferred from saved text",
  choose: "Choose two versions to compare.", chooseGroup: "Choose up to two versions", before: "Before", after: "After", selected: "Selected",
  changes: "Version changes", difference: "Before and after the change", original: "Original comment passage",
  pairCount: (before: string, after: string, n: number) => `${before} → ${after} · ${n} changes`,
  chooseChange: "Choose a change to see its passages and comments.", export: "Export", outputRange: "Maximum text range to export",
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
  tracking: "比較する版", inferred: "保存版の比較から推定",
  choose: "比較する版を2つ選んでください。", chooseGroup: "比較する版を最大2つ選択", before: "比較前", after: "比較後", selected: "選択中",
  changes: "版の変更箇所", difference: "変更箇所の前後", original: "コメントが付いた元の箇所",
  pairCount: (before, after, n) => `${before}→${after}・${n}件`,
  chooseChange: "変更箇所を選ぶと、前後の原文とコメントを表示します。", export: "書き出す", outputRange: "書き出す原文の上限",
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

function reportIcon(type: "comment" | "comments" | "avatar"): SVGSVGElement {
  const namespace = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(namespace, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  svg.classList.add("report-icon");
  const path = document.createElementNS(namespace, "path");
  path.setAttribute("d", type === "avatar" ? "M16 8a4 4 0 1 0-8 0 4 4 0 0 0 8 0ZM4 21a8 6 0 0 1 16 0Z"
    : type === "comment" ? "M4 5h16v11H9l-5 4V5Z" : "M3 3h14v10H8l-5 4V3Zm8 13h5l5 4V8h-2");
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
  kind?: string;
  tag?: string;
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

interface ReviewOptions {
  drafts: readonly ReportDraft[];
  picked: readonly number[];
  project: (picked: number[]) => SavedVersionReport;
  included: (id: string) => boolean;
  setIncluded: (id: string, included: boolean) => void;
  authors: string[];
  outputRange: number;
  setOutputRange: (range: number) => void;
  openVersion: (index: number) => void;
  canOpenVersion: (index: number) => boolean;
  commentComposer?: (id: string) => CommentComposer | undefined;
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
  let picked = [...(review?.picked ?? report.initial)];
  let range = Math.min(state.range ?? DEFAULT_REPORT_RANGE, report.range);
  let selected = review ? report.pairs[0]?.changes.find((change) => change.reviewId === state.selected)
    : report.pairs.find((pair) => pair.before === picked[0] && pair.after === picked[1])?.changes[initialChange ?? 0];
  let original: ReportThread | undefined;
  let originButton: HTMLButtonElement | undefined;

  const rootHeader = el("header");
  const title = el("h1", report.title);
  rootHeader.append(title, el("p", s.tracked, "report-hint"));
  const layout = el("div", undefined, "report-layout");
  const main = el("div", undefined, "report-main");
  const sidebar = el("aside", undefined, "report-block report-comments");
  sidebar.id = "report-comments";
  sidebar.setAttribute("aria-label", s.comments);
  state.expanded ??= [true, true, true];
  const expanded = state.expanded;
  const trace = accordion(s.tracking, "report-trace", () => { expanded[0] = !trace.body.hidden; });
  const list = accordion(s.changes, "report-change-list", () => { expanded[1] = !list.body.hidden; });
  const detail = accordion(s.difference, "report-passages", () => { expanded[2] = !detail.body.hidden; updateRangeVisibility(); });
  main.append(list.section, detail.section);
  layout.append(main, sidebar);
  root.replaceChildren(rootHeader, trace.section, layout);
  root.classList.add("saved-report");

  trace.section.classList.add("report-versions");
  const selectedNames = el("span", undefined, "report-selected-names report-collapsed-summary");
  trace.toggle.append(selectedNames);
  state.table ??= { sort: "modifiedAt", descending: false, widths: [] };
  const table = reportVersionTable(drafts, report.lang, s.tracking, state.table, review || drafts.length > 2 ? (index) => {
    picked = picked.includes(index) ? picked.filter((item) => item !== index) : [...picked.slice(-1), index];
    if (review) report = review.project(picked);
    selected = undefined;
    original = undefined;
    draw();
  } : undefined);
  trace.body.append(table.wrapper);

  const summary = el("span", undefined, "report-hint report-accordion-meta");
  summary.setAttribute("aria-live", "polite");
  const selectedTitle = el("span", undefined, "report-selected-title report-collapsed-summary");
  list.toggle.append(summary, selectedTitle);
  const filters = el("div", undefined, "report-filters");
  const kinds = el("select");
  const kindLabel = el("label", s.changeType);
  kindLabel.append(kinds);
  const tags = el("select");
  const tagLabel = el("label", s.tags);
  tagLabel.append(tags);
  const status = el("p", undefined, "report-hint");
  const results = el("ul", undefined, "report-changes");
  filters.append(kindLabel, tagLabel);
  list.body.append(filters, status, results);

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
  const passages = el("div");
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

  const changeFilters = () => { state.kind = kinds.value; state.tag = tags.value; original = undefined; draw(); };
  kinds.onchange = changeFilters;
  tags.onchange = changeFilters;
  rangeInput.oninput = () => {
    state.range = range = Number(rangeInput.value);
    drawPassages();
  };

  const sorted = () => [...picked].sort((a, b) => a - b);
  const pairFor = () => {
    const [before, after] = sorted();
    return review ? report.pairs[0] : report.pairs.find((pair) => pair.before === before && pair.after === after);
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

  function drawSummary(count: number): void {
    const pair = pairFor();
    const text = selected ? changeTitle(selected) : "";
    selectedTitle.textContent = selectedTitle.title = text;
    summary.textContent = !selected && pair ? s.pairCount(report.drafts[pair.before].label, report.drafts[pair.after].label, count) : "";
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

  function draw(): void {
    rangeMenu.open = false;
    const compared = sorted();
    if (report.title) title.textContent = report.title;
    table.update(compared, s.before, s.after, s.selected);
    selectedNames.replaceChildren(...compared.map((index) => {
      const name = drafts[index].filename ?? drafts[index].label;
      const chip = el("span", name);
      chip.title = name;
      return chip;
    }));
    const pair = pairFor();
    let kind = state.kind ?? "";
    let tag = state.tag ?? "";
    const changes = pair?.changes ?? [];
    const availableKinds = new Map(changes.map((change) => [kindValue(change), changeName(change)]));
    const availableTags = new Set(changes.flatMap((change) => change.threads.flatMap((id) =>
      report.threads[id].messages.flatMap((message) => message.tags))));
    if (!availableKinds.has(kind)) kind = "";
    if (!availableTags.has(tag)) tag = "";
    kinds.replaceChildren(option(s.all, ""), ...[...availableKinds].map(([value, label]) => option(label, value)));
    tags.replaceChildren(option(s.all, ""), ...[...availableTags].map((value) => option(`#${value}`, value)));
    kinds.value = kind;
    tags.value = tag;
    state.kind = kind;
    state.tag = tag;
    const shown = changes.filter((item) => (!kind || kindValue(item) === kind) && (!tag || item.threads.some((id) =>
      report.threads[id].messages.some((message) => message.tags.includes(tag)))));
    if (selected && !shown.includes(selected)) selected = undefined;
    state.selected = selected?.reviewId;
    filters.hidden = !pair || !changes.length;
    status.hidden = Boolean(shown.length && !pair?.gap);
    status.textContent = !pair ? s.choose : pair.unavailable ? s.missing : pair.gap ? s.gap : s.noMatches;
    drawSummary(shown.length);
    results.replaceChildren();
    for (const item of shown) {
      const label = changeTitle(item);
      const row = button("", () => {
        rangeMenu.open = false;
        selected = selected === item ? undefined : item;
        state.selected = selected?.reviewId;
        original = undefined;
        for (const previous of results.querySelectorAll("button")) previous.setAttribute("aria-pressed", String(previous === row && selected === item));
        drawSummary(shown.length);
        drawPassages();
        drawThreads();
      });
      row.setAttribute("aria-pressed", String(item === selected));
      const heading = el("span", label);
      const authors = participants(item);
      if (authors) heading.append(authors);
      row.append(heading, el("small", changeName(item)));
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
      if (review && id) {
        const label = el("label", undefined, "report-include-change");
        const check = el("input");
        check.type = "checkbox";
        check.checked = review.included(id);
        check.setAttribute("aria-label", `${s.export} · ${changeTitle(item)}`);
        check.onchange = () => review.setIncluded(id, check.checked);
        label.append(check, s.export);
        li.append(label);
      }
      results.append(li);
    }
    drawPassages();
    drawThreads();
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
    const chosen = sorted();
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
      if (passages.length) addOpenVersion(column, chosen[side === "before" ? 0 : 1]);
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
    const reportVersions = review ? report.initial : sorted();
    const unlocated = reportVersions.flatMap((id) => report.drafts[id]?.unlocated ?? []);
    sidebar.hidden = state.commentsOpen === false;
    layout.classList.toggle("report-comments-open", !sidebar.hidden);
    commentToggle.setAttribute("aria-expanded", String(!sidebar.hidden));
    if (sidebar.hidden) return;
    sidebar.replaceChildren(el("h2", s.comments));
    if (review?.authors.length) {
      const label = el("label", s.author, "report-my-author");
      const selector = el("select");
      selector.append(option(s.noAuthor, ""), ...review.authors.map((author) => option(author, author)));
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
    if (selected?.reviewId && review?.commentComposer) {
      const compose = review.commentComposer(selected.reviewId);
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
  for (const [index, section] of [trace, list, detail].entries()) section.setExpanded(expanded[index]);
}
