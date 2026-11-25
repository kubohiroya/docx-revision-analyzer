import { commentLinkUrl } from "../../../src/lib/commentReferences";
import { el } from "./extensions";
import {
  DEFAULT_REPORT_RANGE, REPORT_LIMITS, reportWindow, type ReportChange, type ReportPassage,
  type ReportThread, type SavedVersionReport,
} from "../reportData";

const EN = {
  title: "Saved version report", tracking: "Trace across saved versions", inferred: "Inferred from saved text",
  choose: "Choose two drafts to compare.", chooseGroup: "Choose up to two drafts", before: "Before", after: "After", selected: "Selected",
  changes: "Changes between selected drafts", difference: "Before and after the change", original: "Original comment passage",
  gap: "There is a gap in the saved record. The comparison uses the available endpoints.",
  missing: "A selected draft could not be read.", tracked: "Observed changes between saves, not exact edit times.",
  changeType: "Change type", tags: "Comment tag", all: "All", noMatches: "No shared changes match this selection.",
  evidence: "Matching evidence", candidates: (n: number) => `${n} possible sources`,
  limited: "The highlighted text shows matching edges. The middle of the passage was not fully aligned.",
  range: "Text range", ranges: ["About 20 characters", "About 80 characters", "About 200 characters", "About 400 characters", "Full page"],
  estimated: "Estimated page", page: "Page", noPage: "Page boundaries are unavailable. Showing surrounding text.",
  noPassage: "There is no corresponding passage on this side.", noText: "Text excerpts were not shared.",
  paragraph: "Paragraph", sentence: "Sentence", comments: "Word comments", noComments: "No Word comments were shared for this change.",
  unlocated: "Comments without a confirmed location", threadMissing: "This saved draft is unavailable.",
  commentAbsent: "Absent from the latest shared readable draft", commentResolved: "Resolved in this draft",
  showOriginal: "Go to the original passage", closeOriginal: "Back to the comparison",
  noOriginal: "The original range could not be located within the shared text.",
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
  title: "保存版レポート", tracking: "文章の追跡", inferred: "保存版の比較から推定",
  choose: "比較する稿を2つ選んでください。", chooseGroup: "比較する稿を最大2つ選択", before: "比較前", after: "比較後", selected: "選択中",
  changes: "選んだ2稿の変更箇所", difference: "変更箇所の前後", original: "コメントが付いた元の箇所",
  gap: "この区間には保全記録の空白があります。確認できた両端を比較しています。",
  missing: "選んだ稿の一方を読み込めませんでした。", tracked: "保存版の間で観測した変化を示します。個々の編集時刻は分かりません。",
  changeType: "変更の種類", tags: "コメントのタグ", all: "すべて", noMatches: "条件に一致する共有済みの変更箇所はありません。",
  evidence: "対応の手掛かり", candidates: (n) => `元の箇所の候補 ${n} 件`,
  limited: "一致する先頭と末尾を強調しています。中間部の細かな差分は特定できません。",
  range: "表示範囲", ranges: ["約20字", "約80字", "40字×5行相当", "40字×10行相当", "1ページ全体"],
  estimated: "推定ページ", page: "ページ", noPage: "ページ境界を取得できないため、前後の抜粋を表示しています。",
  noPassage: "この側に対応する箇所はありません。", noText: "原文は共有されていません。",
  paragraph: "段落", sentence: "文", comments: "Wordコメント", noComments: "この箇所のWordコメントは共有されていません。",
  unlocated: "位置を特定できないコメント", threadMissing: "元の保存版を読み込めません。",
  commentAbsent: "共有された最新の確認可能な稿では消失", commentResolved: "この稿で解決済み",
  showOriginal: "元の保存版の該当箇所へ", closeOriginal: "前後の比較に戻る",
  noOriginal: "共有された原文の中で元の範囲を特定できません。",
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

export function reportChangeLabel(labels: typeof EN, kind: ReportChange["kind"], moved: boolean): string {
  return moved ? kind === "retained" ? labels.moved : `${labels.moved} · ${labels.kinds[kind]}` : labels.kinds[kind];
}

export function renderSavedVersionReport(
  root: HTMLElement, report: SavedVersionReport, openLink?: (url: string) => void,
): void {
  const s = REPORT_LABELS[report.lang];
  let picked = [...report.initial];
  let range = Math.min(DEFAULT_REPORT_RANGE, report.range);
  let selected: ReportChange | undefined;
  let original: ReportThread | undefined;
  let originButton: HTMLButtonElement | undefined;

  const rootHeader = el("header");
  rootHeader.append(el("h1", s.title), el("p", s.tracked, "report-hint"));
  const layout = el("div", undefined, "report-layout");
  const main = el("div", undefined, "report-main");
  const sidebar = el("aside", undefined, "report-block");
  const trace = el("section", undefined, "report-block");
  const list = el("section", undefined, "report-block");
  const detail = el("section", undefined, "report-block");
  main.append(trace, list, detail);
  layout.append(main, sidebar);
  root.replaceChildren(rootHeader, layout);
  root.classList.add("saved-report");

  const traceHead = el("div", undefined, "report-header");
  traceHead.append(el("h2", s.tracking), el("span", s.inferred, "report-hint"));
  const draftChoices = el("div", undefined, "report-drafts");
  draftChoices.setAttribute("role", "group");
  draftChoices.setAttribute("aria-label", s.chooseGroup);
  trace.append(traceHead, draftChoices);
  const draftButtons = report.drafts.map((draft, index) => {
    const tone = el("span");
    const meta = draft.archivedAt ?? draft.modifiedAt;
    const status = draft.unavailable ? s.missing : meta && `${draft.archivedAt ? s.archived : s.modified} ${displayDate(meta)}`;
    const location = el("small", status);
    const node = button("", () => {
      picked = picked.includes(index) ? picked.filter((item) => item !== index) : [...picked.slice(-1), index];
      selected = undefined;
      original = undefined;
      draw();
    });
    node.append(tone, el("strong", draft.label));
    node.append(location);
    draftChoices.append(node);
    return { node, tone, location, status };
  });

  const listHead = el("div", undefined, "report-header");
  const summary = el("span", undefined, "report-hint");
  summary.setAttribute("aria-live", "polite");
  listHead.append(el("h2", s.changes), summary);
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
  list.append(listHead, filters, status, results);

  const detailHead = el("div", undefined, "report-header");
  const heading = el("h2", s.difference);
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
  const detailBody = el("div");
  detail.append(detailHead, context, detailBody);

  const changeFilters = () => { selected = undefined; original = undefined; draw(); };
  kinds.onchange = changeFilters;
  tags.onchange = changeFilters;
  rangeInput.oninput = () => {
    range = Number(rangeInput.value);
    drawPassages();
  };

  const sorted = () => [...picked].sort((a, b) => a - b);
  const rangeLabel = (p: ReportPassage) => [p.heading, `${s.paragraph} ${p.paragraph}`].filter(Boolean).join(" · ");
  const kindValue = (change: ReportChange) => change.reordered ? `moved:${change.kind}` : change.kind;
  const changeName = (change: ReportChange) => reportChangeLabel(s, change.kind, change.reordered);

  function displayDate(date: string): string {
    const time = new Date(date);
    return Number.isFinite(time.valueOf()) ? time.toLocaleString(report.lang) : "";
  }

  function draw(): void {
    const [before, after] = sorted();
    for (const [index, { node, tone }] of draftButtons.entries()) {
      const chosen = picked.includes(index);
      node.setAttribute("aria-pressed", String(chosen));
      tone.textContent = chosen ? picked.length === 1 ? s.selected : index === before ? s.before : s.after : "";
    }
    const pair = report.pairs.find((item) => item.before === before && item.after === after);
    let kind = kinds.value;
    let tag = tags.value;
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
    const shown = changes.filter((item) => (!kind || kindValue(item) === kind) && (!tag || item.threads.some((id) =>
      report.threads[id].messages.some((message) => message.tags.includes(tag)))));
    if (!selected || !shown.includes(selected)) selected = shown[0];
    filters.hidden = !pair || !changes.length;
    status.hidden = Boolean(shown.length && !pair?.gap);
    status.textContent = !pair ? s.choose : pair.unavailable ? s.missing : pair.gap ? s.gap : s.noMatches;
    summary.textContent = pair ? `${report.drafts[before].label} → ${report.drafts[after].label} · ${shown.length}` : "";
    results.replaceChildren();
    for (const item of shown) {
      const left = item.before[0];
      const right = item.after[0];
      const label = left && right && left.heading !== right.heading
        ? `${rangeLabel(left)} → ${rangeLabel(right)}` : rangeLabel(right ?? left);
      const row = button("", () => {
        selected = item;
        original = undefined;
        for (const previous of results.querySelectorAll("button")) previous.setAttribute("aria-pressed", String(previous === row));
        drawLocations();
        drawPassages();
        drawThreads();
      });
      row.setAttribute("aria-pressed", String(item === selected));
      row.append(el("span", label), el("small", changeName(item)));
      const first = right ?? left;
      const excerpt = first.excerpt;
      if (excerpt) {
        const text = report.clips[excerpt.clip].text;
        const [start, end] = reportWindow(text, excerpt.focus, 44);
        row.append(el("span", (start ? "…" : "") + text.slice(start, end).replace(/\s+/gu, " ") + (end < text.length ? "…" : ""), "report-row-text"));
      }
      const li = el("li");
      li.append(row);
      results.append(li);
    }
    drawLocations();
    drawPassages();
    drawThreads();
  }

  function drawLocations(): void {
    const [before, after] = sorted();
    for (const [index, { location, status }] of draftButtons.entries()) {
      const passage = index === before ? selected?.before[0] : index === after ? selected?.after[0] : undefined;
      location.textContent = passage ? rangeLabel(passage) : status ?? "";
      location.title = passage ? status ?? "" : "";
    }
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
    detail.hidden = !selected && !original;
    if (detail.hidden) return;
    heading.textContent = original ? s.original : s.difference;
    detailHead.replaceChildren(heading);
    context.hidden = Boolean(original) || !report.includeText;
    detailBody.replaceChildren();
    if (original) {
      const back = button(s.closeOriginal, () => {
        original = undefined;
        drawPassages();
        originButton?.focus({ preventScroll: true });
      });
      detailHead.append(back);
      detailBody.append(el("h3", report.drafts[original.draft].label));
      for (const passage of original.original ?? []) detailBody.append(passageView(passage, "original"));
      return;
    }
    const change = selected!;
    const [before, after] = sorted();
    rangeValue.textContent = s.ranges[range];
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
      columns.append(column);
    }
    detailBody.append(columns);
  }

  function showThread(thread: ReportThread, allowTagFilter: boolean): HTMLElement {
    const node = el("article", undefined, "report-thread");
    node.append(el("h3", report.drafts[thread.draft].label));
    const notes = [thread.absent && s.commentAbsent, thread.resolved && s.commentResolved, thread.unavailable && s.threadMissing].filter(Boolean);
    if (notes.length) node.append(el("p", notes.join(" · "), "report-hint"));
    for (const [index, message] of thread.messages.entries()) {
      const post = el("div", undefined, "report-post");
      if (index) post.append(el("p", s.reply, "report-hint"));
      post.append(el("strong", message.author || s.unnamed));
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
        drawPassages();
        detailHead.querySelector("button")?.focus({ preventScroll: true });
        detail.scrollIntoView({ block: "nearest" });
      }, "report-original");
      node.append(show);
    } else {
      node.append(el("p", s.noOriginal, "report-hint"));
    }
    return node;
  }

  function drawThreads(): void {
    const unlocated = sorted().flatMap((id) => report.drafts[id].unlocated);
    sidebar.hidden = !selected && !unlocated.length;
    sidebar.replaceChildren(el("h2", s.comments));
    if (selected?.threads.length) {
      for (const id of selected.threads) sidebar.append(showThread(report.threads[id], true));
    } else if (selected) {
      sidebar.append(el("p", s.noComments, "report-hint"));
    }
    if (unlocated.length) {
      const box = el("details");
      box.append(el("summary", s.unlocated));
      for (const id of unlocated) box.append(showThread(report.threads[id], false));
      sidebar.append(box);
    }
  }

  draw();
}
