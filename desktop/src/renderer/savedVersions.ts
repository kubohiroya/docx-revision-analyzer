import {
  compareSavedVersions, lineageRangeText, traceLineage,
  type LineageKind, type LineageLevel, type LineageLink, type LineageRange,
  type MatchEvidence, type SavedTextVersion, type SavedVersionLineage,
} from "../../../src/lib/lineage";
import { commentReferences } from "../../../src/lib/commentReferences";
import { LineageComments, type VersionThreadOccurrence } from "../lineageComments";
import { SavedReportAnalysis } from "../savedReport";
import type { AppLang, CommentHistoryTarget, SavedLineageVersion } from "../shared";
import { el } from "./extensions";
import { showCommentHistory } from "./commentHistory";
import { savedReportExportLabel, showSavedReportExport } from "./savedReportExport";
import type { FigureSettings } from "./savedReportFigures";
import { STRINGS } from "./strings";

const EN = {
  preview: "Compare saved versions (preview)", tab: "Saved versions", version: "Version",
  select: "Choose DOCX files", folders: "Choose folders", noInputs: "No DOCX files were found.", archive: "Use the open document’s saved record",
  order: "Order the files from earliest to latest, then compare.", compare: "Compare", minimum: "Choose at least two saved versions.",
  working: "Reading saved versions…", missing: "This version could not be read.", noArchive: "There are no archived versions for this document.",
  earlier: "Move earlier", later: "Move later", pair: "Compare versions", unit: "Track by", paragraph: "Paragraph", sentence: "Sentence",
  unchanged: "Show unchanged text", before: "Before", after: "After", relation: "Relationship", reordered: "Order changed",
  interval: "Changes were observed between these versions. The exact edit time and changes made and undone between saves are unknown.",
  gap: "The saved record has a gap here. Only the two available versions can be compared.",
  tracked: "These files also contain Word Track Changes. The comparison uses the visible text in each saved version.",
  noPairs: "No matching changes to display.", evidence: "Matching evidence", candidate: "Possible source",
  sample: (shown: number, total: number) => `Showing ${shown} of ${total} candidates.`,
  limited: "The highlighted beginning and end match. The middle of this long passage has no detailed alignment.",
  timeline: "Related occurrences across versions", archivedAt: "Archived", modifiedAt: "File modified",
  comments: "Word comments on related passages", versionComments: "Word comments by version", commentsVersion: "Comment version",
  noComments: "There are no recorded threads in this version.", commentsUnreadable: "Word comments in this version could not be read.",
  commentAbsent: "Absent in the latest available version", commentResolved: "Resolved in this version", openThread: "Open thread",
  openOriginal: "Go to the original passage", sourceText: "Original saved text", closeSource: "Close original text",
  countThreads: (count: number) => `${count} Word comment thread(s)`,
  kinds: {
    retained: "Unchanged", rewritten: "Rewritten", split: "Split", merged: "Merged", copied: "Exact copy candidate",
    reappeared: "Reappearance candidate", added: "Added", removed: "No longer present", ambiguous: "Uncertain",
  } satisfies Record<LineageKind, string>,
  clues: {
    exact: "Exact text match after whitespace and Unicode normalization", paragraphId: "Word paragraph ID",
    neighbors: "Matching neighboring text", heading: "Same heading", characterOverlap: "Shared character sequences",
    japaneseWords: "Shared Japanese word segments", joined: "Text matches after joining adjacent passages", past: "Identical text in an earlier version",
  } satisfies Record<MatchEvidence, string>,
};
const LABELS: Record<AppLang, typeof EN> = { en: EN, ja: {
  preview: "保存版を比較（プレビュー）", tab: "保存版の比較", version: "保存版",
  select: "DOCXファイルを選ぶ", folders: "フォルダを選ぶ", noInputs: "対象のDOCXファイルが見つかりませんでした。", archive: "開いている文書の保全記録を使う",
  order: "ファイルを古い順に並べてから比較します。", compare: "比較する", minimum: "保存版を2つ以上選んでください。",
  working: "保存版を読み込んでいます…", missing: "この保存版を読み込めませんでした。", noArchive: "この文書の保全済み保存版はありません。",
  earlier: "前へ移す", later: "後ろへ移す", pair: "比較する版", unit: "追跡単位", paragraph: "段落", sentence: "文",
  unchanged: "変更のない箇所も表示", before: "前の版", after: "次の版", relation: "対応", reordered: "並び順が変化",
  interval: "この二つの版の間で確認できた変化を示します。正確な編集時刻や、保存の間に書いて消した内容は分かりません。",
  gap: "この区間には保全記録の空白があります。確認できた両端の版を比較しています。",
  tracked: "Wordの変更履歴を含む版があります。比較には各保存版の変更を反映した本文を使います。",
  noPairs: "表示する対応はありません。", evidence: "対応の手掛かり", candidate: "元の箇所の候補",
  sample: (shown, total) => `候補 ${total} 件のうち ${shown} 件を表示しています。`,
  limited: "長い文章のため、一致する先頭と末尾だけを強調しています。",
  timeline: "版をまたぐ関連箇所", archivedAt: "保管日時", modifiedAt: "ファイル更新日時",
  comments: "関連箇所のWordコメント", versionComments: "保存版ごとのWordコメント", commentsVersion: "コメントを読む版",
  noComments: "この版に記録されたスレッドはありません。", commentsUnreadable: "この版のWordコメントを読み取れませんでした。",
  commentAbsent: "確認できる最新の版では消失", commentResolved: "この版で解決済み", openThread: "スレッドを開く",
  openOriginal: "元の保存版の該当箇所へ", sourceText: "当時の保存版の本文", closeSource: "元の本文を閉じる",
  countThreads: (count) => `Wordコメント ${count} 件`,
  kinds: {
    retained: "継続", rewritten: "書き換え", split: "分割", merged: "結合", copied: "複製候補",
    reappeared: "再出現候補", added: "追加", removed: "消失", ambiguous: "未確定",
  },
  clues: {
    exact: "空白・Unicodeを正規化した文字列が一致", paragraphId: "Wordの段落ID",
    neighbors: "前後の文章が対応", heading: "同じ見出し", characterOverlap: "共通する文字列",
    japaneseWords: "共通する日本語の語区切り", joined: "隣り合う箇所をつなぐと一致", past: "過去の版に同じ文面",
  },
} };

export function savedVersionsLabel(lang: AppLang, inSettings = false): string {
  return inSettings ? LABELS[lang].preview : LABELS[lang].tab;
}

function action(label: string, click: () => void, disabled = false): HTMLButtonElement {
  const button = el("button", label);
  button.disabled = disabled;
  button.onclick = click;
  return button;
}

function option(label: string, value: string): HTMLOptionElement {
  const item = el("option", label);
  item.value = value;
  return item;
}

function timestamp(...values: (string | undefined)[]): number {
  for (const value of values) {
    const parsed = value ? Date.parse(value) : NaN;
    if (Number.isFinite(parsed)) return parsed;
  }
  return Number.MAX_SAFE_INTEGER;
}

function compareVersionTime(left: SavedLineageVersion, right: SavedLineageVersion): number {
  return timestamp(left.archivedAt, left.sourceModified, left.fileCreatedAt) -
    timestamp(right.archivedAt, right.sourceModified, right.fileCreatedAt) ||
    (left.number ?? 0) - (right.number ?? 0) ||
    timestamp(left.fileCreatedAt) - timestamp(right.fileCreatedAt) ||
    left.label.localeCompare(right.label, "en", { numeric: true }) ||
    (left.sha256 ?? "").localeCompare(right.sha256 ?? "") ||
    (left.sourcePathHash ?? "").localeCompare(right.sourcePathHash ?? "");
}

export class SavedVersionsPane {
  private versions: SavedLineageVersion[] = [];
  private result?: SavedVersionLineage;
  private comments?: LineageComments;
  private commentVersion?: number;
  private versionCommentsOpen = false;
  private original?: VersionThreadOccurrence;
  private selected?: LineageLink;
  private pair = 0;
  private level: LineageLevel = "paragraph";
  private showUnchanged = false;
  private busy = false;
  private notice: "minimum" | "noArchive" | "noInputs" | undefined;
  private error = "";
  private lang: AppLang = "en";
  private current?: CommentHistoryTarget;

  constructor(
    private pane: HTMLElement,
    private figureSettings: FigureSettings,
    private branchTimelineEnabled: () => boolean,
  ) {}

  private get labels() { return LABELS[this.lang]; }

  private name(index: number): string {
    const v = this.versions[index];
    return v.source === "archive"
      ? `${this.labels.version} ${v.number ?? index + 1}`
      : `${index + 1}. ${v.label}`;
  }

  private date(v: SavedTextVersion): string {
    const iso = v.archivedAt ?? v.sourceModified;
    if (!iso) return "";
    return `${v.archivedAt ? this.labels.archivedAt : this.labels.modifiedAt} ${new Date(iso).toLocaleString(this.lang)}`;
  }

  loadPaths(paths: string[]): Promise<void> {
    return this.load(() => window.app.loadSavedVersionPaths(paths), { review: true });
  }

  private async load(
    read: () => Promise<SavedLineageVersion[] | null>, { fromArchive = false, review = false } = {},
  ): Promise<void> {
    if (this.busy) return;
    this.busy = true;
    this.notice = undefined;
    this.error = "";
    this.draw();
    let loaded = false;
    try {
      const items = await read();
      if (!items) return;
      if (!items.length && fromArchive) {
        this.notice = "noArchive";
        return;
      }
      this.versions = items;
      this.result = undefined;
      this.comments = undefined;
      this.commentVersion = undefined;
      this.original = undefined;
      this.selected = undefined;
      this.pair = 0;
      if (!items.length) this.notice = "noInputs";
      else if (items.length < 2) this.notice = "minimum";
      loaded = items.length > 0;
    } catch (error) {
      this.error = String(error);
    } finally {
      this.busy = false;
      this.draw();
    }
    if (loaded && review) this.openReview();
  }

  private openReview(): void {
    const ordered = this.versions.map((version, index) => ({ version, index }));
    ordered.sort((a, b) => compareVersionTime(a.version, b.version));
    const versions = ordered.map(({ version }) => version);
    const lineage = compareSavedVersions(versions, { japaneseWordHints: true });
    void showSavedReportExport(
      new SavedReportAnalysis(versions, lineage, this.level), versions,
      ordered.map(({ index }) => this.name(index)), this.lang, this.figureSettings, this.branchTimelineEnabled(),
      (index, version) => {
        this.versions[ordered[index].index] = version;
        if (this.result) {
          this.result = { ...this.result, versions: [...this.versions] };
          this.comments = new LineageComments(this.versions, this.result);
        }
        this.draw();
      },
    );
  }

  private reorder(index: number, delta: -1 | 1): void {
    const destination = index + delta;
    [this.versions[index], this.versions[destination]] = [this.versions[destination], this.versions[index]];
    this.result = undefined;
    this.comments = undefined;
    this.commentVersion = undefined;
    this.original = undefined;
    this.selected = undefined;
    this.pair = 0;
    this.draw();
    const row = this.pane.querySelectorAll(".lineage-versions > li")[destination];
    (row?.querySelector<HTMLButtonElement>("button:not(:disabled)"))?.focus();
  }

  private compare(): void {
    try {
      this.error = "";
      this.result = compareSavedVersions([...this.versions], { japaneseWordHints: true });
      this.comments = new LineageComments(this.versions, this.result);
      this.selected = undefined;
    } catch (error) {
      this.error = String(error);
    }
    this.draw();
    this.pane.querySelector<HTMLSelectElement>("#lineage-pair")?.focus();
  }

  render(lang: AppLang, current?: CommentHistoryTarget): void {
    this.lang = lang;
    this.current = current;
    this.draw();
  }

  private draw(): void {
    const s = this.labels;
    this.pane.replaceChildren();
    const controls = el("div", undefined, "lineage-controls");
    const current = this.current;
    controls.append(
      action(s.select, () => void this.load(() => window.app.chooseSavedVersions(), { review: true }), this.busy),
      action(s.folders, () => void this.load(() => window.app.chooseSavedVersionFolders(), { review: true }), this.busy),
      action(s.archive, () => {
        if (current) void this.load(() => window.app.watch.savedVersionsForComparison(current), { fromArchive: true });
      }, this.busy || !current),
    );
    this.pane.append(controls);
    if (this.busy) this.pane.append(el("p", s.working, "hint"));
    if (this.notice) this.pane.append(el("p", s[this.notice], "hint"));
    if (this.error) this.pane.append(el("p", this.error, "error"));
    if (!this.versions.length) return;
    const manual = this.versions[0].source === "selected";
    if (manual) this.pane.append(el("p", s.order, "hint"));
    const versions = el("ol", undefined, "lineage-versions");
    for (const [index, v] of this.versions.entries()) {
      const row = el("li");
      row.append(el("span", this.name(index)));
      const date = this.date(v);
      if (date) row.append(el("span", date, "hint"));
      if (v.unavailable) row.append(el("span", s.missing, "error"));
      if (manual) {
        const up = action("↑", () => this.reorder(index, -1), index === 0 || this.busy);
        const down = action("↓", () => this.reorder(index, 1), index === this.versions.length - 1 || this.busy);
        up.setAttribute("aria-label", `${this.name(index)} — ${s.earlier}`);
        down.setAttribute("aria-label", `${this.name(index)} — ${s.later}`);
        row.append(up, down);
      }
      versions.append(row);
    }
    this.pane.append(versions, action(s.compare, () => this.compare(), this.versions.length < 2 || this.busy));
    if (this.original) this.drawOriginal(this.original);
    if (this.result) this.drawResult(this.result);
  }

  private kind(link: LineageLink): string {
    const s = this.labels;
    return link.reordered ? s.kinds[link.kind] + " · " + s.reordered : s.kinds[link.kind];
  }

  private brief(ranges: LineageRange[]): string {
    const text = ranges.map((range) => lineageRangeText(this.versions, range)).join(" / ");
    const chars = Array.from(text);
    return chars.length > 72 ? chars.slice(0, 72).join("") + "…" : text;
  }

  private drawResult(result: SavedVersionLineage): void {
    const s = this.labels;
    this.pane.append(el("p", s.interval, "hint"));
    if (result.versions.some((v) => !v.unavailable && v.trackedChanges)) this.pane.append(el("p", s.tracked, "hint"));
    const controls = el("div", undefined, "lineage-controls");
    const pairLabel = el("label", s.pair);
    const pairs = el("select");
    pairs.id = "lineage-pair";
    for (const [i, p] of result.comparisons.entries()) pairs.append(option(`${this.name(p.before)} → ${this.name(p.after)}`, String(i)));
    pairs.value = String(this.pair);
    pairs.onchange = () => {
      this.pair = Number(pairs.value);
      this.selected = undefined;
      this.draw();
      this.pane.querySelector<HTMLSelectElement>("#lineage-pair")?.focus();
    };
    pairLabel.append(pairs);
    const levelLabel = el("label", s.unit);
    const levels = el("select");
    levels.id = "lineage-level";
    levels.append(option(s.paragraph, "paragraph"), option(s.sentence, "sentence"));
    levels.value = this.level;
    levels.onchange = () => {
      this.level = levels.value === "sentence" ? "sentence" : "paragraph";
      this.selected = undefined;
      this.draw();
      this.pane.querySelector<HTMLSelectElement>("#lineage-level")?.focus();
    };
    levelLabel.append(levels);
    const showLabel = el("label", undefined, "check");
    const show = el("input");
    show.id = "lineage-unchanged";
    show.type = "checkbox";
    show.checked = this.showUnchanged;
    show.onchange = () => {
      this.showUnchanged = show.checked;
      this.selected = undefined;
      this.draw();
      this.pane.querySelector<HTMLInputElement>("#lineage-unchanged")?.focus();
    };
    showLabel.append(show, s.unchanged);
    const pair = result.comparisons[this.pair];
    controls.append(pairLabel, levelLabel, showLabel, action(savedReportExportLabel(this.lang), () => this.openReview(), this.busy));
    this.pane.append(controls);
    this.drawVersionComments(pair.before, pair.after);
    if (pair.unavailable) {
      this.pane.append(el("p", s.missing, "hint"));
      return;
    }
    if (pair.recordingGap) this.pane.append(el("p", s.gap, "hint"));
    const changes = pair[this.level].filter((link) =>
      this.showUnchanged || link.kind !== "retained" || link.reordered || this.threadsFor(link).length > 0);
    if (!changes.length) {
      this.pane.append(el("p", s.noPairs, "hint"));
      return;
    }
    const heading = el("div", undefined, "lineage-row lineage-row-heading");
    heading.append(el("span", s.before), el("span", s.relation), el("span", s.after));
    const list = el("div", undefined, "lineage-list");
    const detail = el("section", undefined, "lineage-detail");
    if (!this.selected || !changes.includes(this.selected)) this.selected = changes[0];
    for (const change of changes) {
      const row = el("button", undefined, "lineage-row");
      row.setAttribute("aria-pressed", String(this.selected === change));
      const relationship = el("span");
      relationship.append(el("strong", this.kind(change)));
      const count = this.threadsFor(change).length;
      if (count) relationship.append(el("small", s.countThreads(count), "lineage-thread-count"));
      row.append(el("span", this.brief(change.from)), relationship, el("span", this.brief(change.to)));
      row.onclick = () => {
        this.selected = change;
        for (const button of list.querySelectorAll("button")) button.setAttribute("aria-pressed", String(button === row));
        this.drawDetail(detail, change, result);
      };
      list.append(row);
    }
    this.pane.append(heading, list, detail);
    this.drawDetail(detail, this.selected, result);
  }

  private drawDetail(detail: HTMLElement, change: LineageLink, result: SavedVersionLineage): void {
    const s = this.labels;
    detail.replaceChildren(el("h2", this.kind(change)));
    if (change.evidence.length) detail.append(el("p", `${s.evidence} · ${change.evidence.map((e) => s.clues[e]).join(" / ")}`, "hint"));
    if (change.alignmentPartial) detail.append(el("p", s.limited, "hint"));
    const passages = el("div", undefined, "lineage-passages");
    for (const side of ["from", "to"] as const) {
      const column = el("section");
      column.append(el("h3", side === "from" ? change.kind === "ambiguous" ? s.candidate : s.before : s.after));
      if (side === "from" && change.candidateCount) column.append(el("p", s.sample(change.from.length, change.candidateCount), "hint"));
      for (const range of change[side]) {
        column.append(el("p", `${this.name(range.version)} · ${s.paragraph} ${range.paragraph + 1}`, "hint"));
        const passage = el("p", undefined, "lineage-passage");
        const text = lineageRangeText(this.versions, range);
        const matches = change.unchanged.map((p) => p[side]).filter((r) =>
          r.paragraph === range.paragraph && r.version === range.version && r.start >= range.start && r.end <= range.end);
        let start = range.start;
        for (const match of matches) {
          if (match.start > start) {
            const gap = text.slice(start - range.start, match.start - range.start);
            passage.append(change.alignmentPartial ? gap : el("mark", gap));
          }
          const matched = text.slice(match.start - range.start, match.end - range.start);
          passage.append(change.alignmentPartial ? el("mark", matched) : matched);
          start = match.end;
        }
        if (start < range.end) {
          const tail = text.slice(start - range.start);
          passage.append(change.kind === "ambiguous" || change.alignmentPartial ? tail : el("mark", tail));
        }
        column.append(passage);
      }
      passages.append(column);
    }
    detail.append(passages);
    const threads = this.threadsFor(change);
    if (threads.length) detail.append(el("h3", s.comments), this.threadList(threads));
    const anchor = change.to[0] ?? change.from[0];
    const lineage = anchor ? traceLineage(result, this.level, anchor) : [];
    if (lineage.length > 1) {
      const list = el("ul", undefined, "lineage-timeline");
      for (const entry of lineage) {
        const prior = entry.from[0];
        const next = entry.to[0];
        const labels = [prior && this.name(prior.version), this.kind(entry), next && this.name(next.version)].filter(Boolean);
        const item = el("li");
        const visibleText = `${labels.join(" · ")} — ${this.brief(entry.to.length ? entry.to : entry.from)}`;
        const comparison = result.comparisons.findIndex((p) => p[this.level].includes(entry));
        item.append(action(visibleText, () => {
          this.pair = comparison;
          this.selected = entry;
          if (entry.kind === "retained" && !entry.reordered) this.showUnchanged = true;
          this.draw();
          this.pane.querySelector<HTMLButtonElement>('.lineage-row[aria-pressed="true"]')?.focus();
        }));
        list.append(item);
      }
      detail.append(el("h3", s.timeline), list);
    }
  }

  private threadsFor(link: LineageLink): VersionThreadOccurrence[] {
    return this.comments?.forRelationship(link, this.level) ?? [];
  }

  private drawVersionComments(before: number, after: number): void {
    const s = this.labels;
    const box = el("details", undefined, "lineage-version-comments");
    box.open = this.versionCommentsOpen;
    box.ontoggle = () => { this.versionCommentsOpen = box.open; };
    box.append(el("summary", s.versionComments));
    const label = el("label", s.commentsVersion);
    const versions = el("select");
    versions.id = "lineage-comment-version";
    for (const [index, version] of this.versions.entries()) {
      const name = version.threads.length ? `${this.name(index)} · ${s.countThreads(version.threads.length)}` : this.name(index);
      versions.append(option(name, String(index)));
    }
    const current = this.commentVersion ?? (this.versions[after].threads.length ? after : before);
    versions.value = String(current);
    versions.onchange = () => {
      this.commentVersion = Number(versions.value);
      this.draw();
      this.pane.querySelector<HTMLSelectElement>("#lineage-comment-version")?.focus();
    };
    label.append(versions);
    box.append(label);
    const selected = this.versions[current];
    if (selected.commentsUnavailable) box.append(el("p", s.commentsUnreadable, "hint"));
    else if (selected.threads.length) box.append(this.threadList(selected.threads.map((thread) => ({ version: current, thread }))));
    else box.append(el("p", s.noComments, "hint"));
    this.pane.append(box);
  }

  private threadList(occurrences: VersionThreadOccurrence[]): HTMLElement {
    const s = this.labels;
    const list = el("ul", undefined, "lineage-threads");
    for (const entry of occurrences) {
      const messages = entry.thread.detail.comments;
      const author = messages[0].author ?? "";
      const tags = [...new Set(messages.flatMap((comment) => commentReferences(comment.text, comment.links).tags))];
      const states = this.threadStates(entry);
      const text = Array.from(messages[0].text.replace(/\s+/g, " ").trim());
      const summary = text.length > 100 ? text.slice(0, 100).join("") + "…" : text.join("");
      const item = el("li");
      const byline = [this.name(entry.version), author, ...states].filter(Boolean).join(" · ");
      item.append(el("div", byline, "hint"), el("p", summary, "comment-text"));
      if (tags.length) item.append(el("p", tags.map((tag) => `#${tag}`).join(" "), "hint"));
      item.append(action(s.openThread, () => void this.openThread(entry)));
      list.append(item);
    }
    return list;
  }

  private threadStates(entry: VersionThreadOccurrence): string[] {
    const states: string[] = [];
    if (entry.thread.detail.comments[0].done) states.push(this.labels.commentResolved);
    if (this.comments?.absentInLatest(entry)) states.push(this.labels.commentAbsent);
    return states;
  }

  private async openThread(entry: VersionThreadOccurrence): Promise<void> {
    const version = this.versions[entry.version];
    const date = (iso: string) => new Date(iso).toLocaleString(this.lang);
    const versionLabel = [this.name(entry.version), this.date(version), ...this.threadStates(entry)].filter(Boolean).join(" · ");
    await showCommentHistory(entry.thread.detail, STRINGS[this.lang], date, (url) => {
      void window.app.watch.openCommentLink(url).catch((error) => {
        this.error = String(error);
        this.draw();
      });
    }, {
      versionLabel,
      original: entry.thread.ranges.length && !version.unavailable ? {
        label: this.labels.openOriginal,
        show: () => {
          if (this.versions[entry.version] !== version) return;
          this.original = entry;
          this.draw();
          const original = this.pane.querySelector<HTMLElement>(".lineage-original");
          original?.focus({ preventScroll: true });
          original?.scrollIntoView({ block: "start" });
        },
      } : undefined,
    });
  }

  private drawOriginal(entry: VersionThreadOccurrence): void {
    const version = this.versions[entry.version];
    if (version.unavailable) return;
    const s = this.labels;
    const original = el("section", undefined, "lineage-original");
    original.tabIndex = -1;
    original.append(el("h2", `${s.sourceText} · ${this.name(entry.version)}`));
    const marks = new Map(entry.thread.ranges.map((range) => [range.paragraph, range]));
    for (const [index, p] of version.paragraphs.entries()) {
      const mark = marks.get(p.index);
      if (!mark && !marks.has(version.paragraphs[index - 1]?.index) && !marks.has(version.paragraphs[index + 1]?.index)) continue;
      const paragraph = el("p", undefined, "lineage-passage");
      if (!mark) paragraph.append(p.text);
      else {
        const highlight = el("mark", p.text.slice(mark.start, mark.end) || "│");
        if (mark.start === mark.end) highlight.setAttribute("aria-label", STRINGS[this.lang].commentHistoryPoint);
        paragraph.append(p.text.slice(0, mark.start), highlight, p.text.slice(mark.end));
      }
      original.append(el("div", `${s.paragraph} ${p.index + 1}`, "hint"), paragraph);
    }
    original.append(action(s.closeSource, () => {
      this.original = undefined;
      this.pane.querySelector<HTMLElement>('.lineage-row[aria-pressed="true"]')?.focus();
      original.remove();
    }));
    this.pane.append(original);
  }
}
