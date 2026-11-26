import { commentReferences } from "../../../src/lib/commentReferences";
import { projectHistoricalAnchor, type HistoricalMatch } from "../../../src/lib/historicalAnchors";
import { compareSavedVersions, type SavedParagraphRange } from "../../../src/lib/lineage";
import { historyEntries, historyMatch, searchKey, type HistoryEntry } from "../historySearch";
import type { AppLang, SavedLineageVersion } from "../shared";
import { el } from "./extensions";

const labels = {
  en: {
    history: "Saved history", title: "Search saved history", search: "Text or quotation", all: "All", kind: "Kind", author: "Author", tag: "Tag", dateBasis: "Filter by date",
    version: "Version", from: "From", to: "Through", state: "State", resolved: "Resolved",
    absent: "No longer observed", unknownDate: "Unknown dates", open: "Show source", file: "Open source file",
    target: "Find corresponding position in", choose: "Choose a version", previous: "Previous", next: "Next",
    quote: "Recorded quotation", noPosition: "The original range could not be verified.", matched: "matches",
    revisionCoverage: "Change history read", commentCoverage: "Comment files read", unavailable: "Unavailable", pending: "Pending", changed: "Changed since comparison", indexing: "Reading versions…",
    excluded: "Outside index scope", failed: "Unable to read change history", disabled: "Indexing disabled",
    tables: "tables", headersFooters: "headers and footers", notes: "footnotes and endnotes", textBoxes: "text boxes", formatting: "formatting changes",
    limited: "Positions and change history cover body paragraphs and headings. Text from unavailable versions may be missing.",
    notOrdered: "Choose a branch or compare versions in a confirmed order to find a corresponding position.",
    absentAt: "No longer observed in", range: "Corresponding range", partial: "Part of the original range",
    boundary: "Estimated boundary", related: "Related paragraph", ambiguous: "Multiple possible locations",
    deleted: "No corresponding passage found; it may have been deleted", unknown: "Position could not be determined",
    gap: "There is a gap or an unreadable version on this path.", close: "Back to comparison",
    ins: "Insertion", del: "Deletion", moveTo: "Move destination", moveFrom: "Move source", comment: "Comment", reply: "Reply",
    archivedAt: "Archived", sourceModified: "Version date", fetchedAt: "Retrieved", date: "Word date",
  },
  ja: {
    history: "保存版の履歴", title: "保存版の履歴を検索", search: "本文・引用の語句", all: "すべて", kind: "種類", author: "作成者", tag: "タグ", dateBasis: "対象にする日時",
    version: "版", from: "開始", to: "終了", state: "状態", resolved: "解決済み",
    absent: "後続版では観測されない", unknownDate: "日時が不明", open: "出典を表示", file: "元のファイルを開く",
    target: "対応箇所を探す版", choose: "版を選ぶ", previous: "前へ", next: "次へ",
    quote: "記録された引用", noPosition: "元の範囲を確認できませんでした。", matched: "件",
    revisionCoverage: "変更履歴を読めた版", commentCoverage: "コメントを読めた版", unavailable: "読取不能", pending: "未処理", changed: "比較後に変更", indexing: "保存版を読み込んでいます…",
    excluded: "索引の対象外", failed: "変更履歴の読取失敗", disabled: "索引化が無効",
    tables: "表", headersFooters: "ヘッダー・フッター", notes: "脚注・文末脚注", textBoxes: "テキストボックス", formatting: "書式変更",
    limited: "位置照合と変更履歴の対象は本文と見出しです。読み取れない版の内容は検索に出ない場合があります。",
    notOrdered: "対応箇所を探すには、経路を選ぶか、順序を確認した版を比較してください。",
    absentAt: "観測されなくなった版", range: "対応する範囲", partial: "一部が対応する範囲",
    boundary: "推定した境界", related: "関連する段落", ambiguous: "複数の対応候補",
    deleted: "対応箇所なし（削除された可能性）", unknown: "位置を特定できませんでした",
    gap: "この経路には記録の空白または読めない版があります。", close: "比較に戻る",
    ins: "挿入", del: "削除", moveTo: "移動先", moveFrom: "移動元", comment: "コメント", reply: "返信",
    archivedAt: "保管", sourceModified: "版の日時", fetchedAt: "取得", date: "Word内の日時",
  },
};

export function historySearchLabel(lang: AppLang, search = true): string { return search ? labels[lang].title : labels[lang].history; }

function button(text: string, onClick: () => void): HTMLButtonElement {
  const action = el("button", text);
  action.type = "button";
  action.onclick = onClick;
  return action;
}

function marked(text: string, start: number, end: number) {
  const passage = el("p", undefined, "lineage-passage");
  passage.append(text.slice(0, start), el("mark", text.slice(start, end) || "│"), text.slice(end));
  return passage;
}

export function historySearchPanel(
  lang: AppLang, versions: readonly SavedLineageVersion[], names: readonly string[],
  ordered: boolean, anchors: boolean, search: boolean, signal: AbortSignal, dismiss: () => void,
): HTMLElement {
  const s = labels[lang];
  const panel = el("section", undefined, "report-export history-search");
  const header = el("div", undefined, "report-export-actions");
  header.append(el("h2", historySearchLabel(lang, search)), button(s.close, dismiss));
  const controls = el("div", undefined, "lineage-controls");
  const query = el("input");
  query.type = "search";
  function field(label: string, control: HTMLInputElement | HTMLSelectElement): void {
    const container = el("label", label + " ");
    container.append(control);
    controls.append(container);
  }
  function select(label: string, values: readonly [string, string][]) {
    const control = el("select");
    control.add(new Option(s.all, ""));
    for (const [value, text] of values) control.add(new Option(text, value));
    field(label, control);
    return control;
  }
  field(s.search, query);
  const kinds = ["ins", "del", "moveTo", "moveFrom", "comment", "reply"] as const;
  const kind = select(s.kind, kinds.map(key => [key, s[key]]));
  const author = el("input"); field(s.author, author);
  const tag = el("input"); field(s.tag, tag);
  const version = select(s.version, names.map((n, i) => [String(i), n]));
  const state = select(s.state, [["resolved", s.resolved], ["absent", s.absent]]);
  const dateBasis = el("select");
  const dateFields = ["sourceModified", "date", "archivedAt", "fetchedAt"] as const;
  for (const key of dateFields) dateBasis.add(new Option(s[key], key));
  field(s.dateBasis, dateBasis);
  const from = el("input"); from.type = "date"; field(s.from, from);
  const to = el("input"); to.type = "date"; field(s.to, to);
  const scope = el("details");
  const coverage = names.map(name => el("li", `${name} · ${s.pending}`));
  const scopeRows = el("ul");
  scopeRows.append(...coverage);
  scope.append(el("summary", `${s.version} (${names.length})`), scopeRows);
  const status = el("p", "", "hint");
  status.setAttribute("role", "status");
  const list = el("ol", undefined, "history-results");
  const paging = el("div", undefined, "lineage-controls");
  const detail = el("section", undefined, "lineage-original");
  const choices = el("div", undefined, "lineage-controls");
  const projected = el("section", undefined, "lineage-original");
  const history: HistoryEntry[] = [];
  const calendar = new Map<string, string>();
  function dateFor(entry: HistoryEntry): string {
    const key = dateBasis.value as (typeof dateFields)[number];
    const iso = key === "date" ? entry.date : versions[entry.version][key];
    if (!iso) return "";
    const cached = calendar.get(iso);
    if (cached !== undefined) return cached;
    const time = Date.parse(iso);
    if (!Number.isFinite(time)) return "";
    const date = new Date(time);
    const value = [date.getFullYear(), String(date.getMonth() + 1).padStart(2, "0"), String(date.getDate()).padStart(2, "0")].join("-");
    calendar.set(iso, value);
    return value;
  }
  let read = 0;
  let commentsRead = 0;
  let processed = 0;
  let page = 0;
  const pageSize = 40;
  let lineage: ReturnType<typeof compareSavedVersions> | undefined;
  panel.append(header, scope, controls, el("p", s.limited, "hint"), status, list, paging, detail, choices, projected);

  function addRanges(parent: HTMLElement, index: number, ranges: readonly SavedParagraphRange[], highlight = true): boolean {
    const v = versions[index];
    if (!v || v.unavailable) return false;
    let found = false;
    for (const range of ranges) {
      const para = v.paragraphs.find(p => p.index === range.paragraph);
      if (!para) continue;
      const start = Math.max(0, range.start - 100);
      const end = Math.min(para.text.length, range.end + 100);
      parent.append(highlight ? marked(para.text.slice(start, end), range.start - start, range.end - start)
        : el("p", para.text.slice(start, end), "lineage-passage"));
      found = true;
    }
    return found;
  }

  function displayEntry(entry: HistoryEntry): void {
    const source = versions[entry.version];
    detail.replaceChildren(el("h3", `${names[entry.version]} · ${s[entry.kind]}`));
    choices.replaceChildren();
    projected.replaceChildren();
    function byline(values: [string, string | undefined][]): void {
      const text = values.filter(([, value]) => value).map(([label, value]) => `${label} ${formatDate(value!)}`).join(" · ");
      if (text) detail.append(el("p", text, "hint"));
    }
    byline([[s.date, entry.date], [s.sourceModified, source.sourceModified], [s.archivedAt, source.archivedAt], [s.fetchedAt, source.fetchedAt]]);
    if (entry.absentAt !== undefined) detail.append(el("p", `${s.absentAt} ${names[entry.absentAt]}`, "hint"));
    if (entry.thread) {
      for (const comment of entry.thread.detail.comments) {
        detail.append(el("p", [comment.author, formatDate(comment.dateUtc ?? comment.date ?? "")].filter(Boolean).join(" · "), "hint"));
        detail.append(el("p", comment.text, "lineage-passage"));
        for (const link of commentReferences(comment.text, comment.links).links) {
          detail.append(button(link.label || link.url, () => void window.app.watch.openCommentLink(link.url)));
        }
      }
    } else detail.append(el("p", entry.text, "lineage-passage"));
    if (!addRanges(detail, entry.version, entry.ranges)) {
      detail.append(el("p", s.noPosition, "hint"));
      if (entry.quote) detail.append(el("blockquote", `${s.quote} · ${entry.quote}`));
    }
    if (source.openId) detail.append(button(s.file, () => void window.app.openSavedVersion(source.openId!).catch(() => {
      detail.append(el("p", s.noPosition, "error"));
    })));
    if (anchors && entry.version + 1 < versions.length) {
      if (!ordered) choices.append(el("p", s.notOrdered, "hint"));
      else {
        const select = el("select");
        select.add(new Option(s.choose, ""));
        for (let i = entry.version + 1; i < versions.length; i++) select.add(new Option(names[i], String(i)));
        const label = el("label", s.target + " ");
        label.append(select);
        choices.append(label);
        select.onchange = () => {
          projected.replaceChildren();
          if (!select.value) return;
          lineage ??= compareSavedVersions(versions, { japaneseWordHints: true });
          const target = Number(select.value);
          const match = projectHistoricalAnchor(lineage, entry.version, target, entry.ranges);
          showMatch(match, target);
        };
      }
    }
    detail.scrollIntoView({ block: "nearest" });
  }

  function showMatch(match: HistoricalMatch, target: number): void {
    projected.append(el("h3", names[target]), el("p", s[match.kind]));
    if (match.kind === "unknown" && match.reason === "gap") projected.append(el("p", s.gap, "hint"));
    if (match.kind === "ambiguous") {
      const byVersion = new Map<number, SavedParagraphRange[]>();
      for (const candidate of match.ranges) {
        const ranges = byVersion.get(candidate.version) ?? [];
        ranges.push(candidate);
        byVersion.set(candidate.version, ranges);
      }
      for (const [version, ranges] of byVersion) {
        projected.append(el("h4", names[version]));
        addRanges(projected, version, ranges, false);
      }
    } else addRanges(projected, target, match.ranges, match.kind !== "related");
  }

  function formatDate(iso: string): string {
    const time = Date.parse(iso);
    return Number.isFinite(time) ? new Date(time).toLocaleString(lang) : iso;
  }

  function render(): void {
    if (signal.aborted) return;
    const needle = searchKey(query.value.trim());
    const authorNeedle = searchKey(author.value.trim());
    const tagNeedle = searchKey(tag.value.trim().replace(/^#/, ""));
    let unknown = 0;
    const filtered = history.filter(entry => {
      if ((kind.value && entry.kind !== kind.value) ||
          (version.value && entry.version !== Number(version.value)) ||
          (authorNeedle && !searchKey(entry.author ?? "").includes(authorNeedle)) ||
          (tagNeedle && !entry.tags.some(t => searchKey(t) === tagNeedle)) ||
          (state.value === "resolved" && !entry.resolved) ||
          (state.value === "absent" && entry.absentAt === undefined) ||
          (needle && !searchKey(entry.text).includes(needle) && !searchKey(entry.quote ?? "").includes(needle))) return false;
      const date = dateFor(entry);
      if (!date) unknown++;
      return (!from.value || (date !== "" && date >= from.value)) && (!to.value || (date !== "" && date <= to.value));
    });
    page = Math.min(page, Math.max(0, Math.ceil(filtered.length / pageSize) - 1));
    status.textContent = [
      `${filtered.length} ${s.matched}`, `${s.revisionCoverage} ${read}/${versions.length}`,
      `${s.commentCoverage} ${commentsRead}/${versions.length}`, unknown ? `${s.unknownDate} ${unknown}` : "",
      processed < versions.length ? s.indexing : "",
    ].filter(Boolean).join(" · ");
    list.replaceChildren();
    for (const entry of filtered.slice(page * pageSize, (page + 1) * pageSize)) {
      const row = el("li");
      const desc = [names[entry.version], s[entry.kind], entry.author, entry.resolved ? s.resolved : undefined,
        entry.absentAt !== undefined ? s.absent : undefined].filter(Boolean).join(" · ");
      const heading = button(desc, () => displayEntry(entry));
      heading.setAttribute("aria-label", `${s.open} · ${desc}`);
      row.append(heading);
      const match = historyMatch(entry.text, needle);
      const quotation = match ? undefined : historyMatch(entry.quote ?? "", needle);
      if (match || quotation) {
        const hit = match ?? quotation!;
        const text = match ? entry.text : entry.quote!;
        const start = Math.max(0, hit.start - 60);
        row.append(marked(text.slice(start, hit.end + 100), hit.start - start, hit.end - start));
      } else row.append(el("p", entry.text.slice(0, 160), "lineage-passage"));
      list.append(row);
    }
    const back = button(s.previous, () => { page--; render(); });
    back.disabled = page === 0;
    const next = button(s.next, () => { page++; render(); });
    next.disabled = (page + 1) * pageSize >= filtered.length;
    paging.replaceChildren(back, next);
  }
  controls.hidden = !search;
  controls.oninput = () => {
    page = 0;
    detail.replaceChildren(); choices.replaceChildren(); projected.replaceChildren();
    render();
  };
  render();

  void (async () => {
    const revisionRead = new Set<number>();
    for (const [index, version] of versions.entries()) {
      const result = version.openId ? await window.app.readSavedVersionRevisions(version.openId).catch(() => undefined) : undefined;
      if (signal.aborted) return;
      if (result?.status === "read") { read++; revisionRead.add(index); }
      const usable = result?.status !== "changed" && (version.source !== "cloud" || result?.status === "read");
      const hasComments = usable && !version.unavailable && !version.commentsUnavailable;
      if (hasComments) commentsRead++;
      const revisionStatus = result?.status ?? "unavailable";
      coverage[index].textContent = [
        names[index],
        revisionStatus === "read" ? s.revisionCoverage : s[revisionStatus],
        result?.status === "read" && result.excluded.length
          ? `${s.excluded} (${result.excluded.map(region => s[region]).join(", ")})` : "",
        hasComments ? s.commentCoverage : `${s.commentCoverage}: ${s.unavailable}`,
      ].filter(Boolean).join(" · ");
      if (usable) {
        for (const entry of historyEntries(version, index, result?.records ?? [])) history.push(entry);
      }
      processed++;
      render();
    }
    if (ordered) {
      annotateDisappearance(versions, history, revisionRead);
      render();
    }
  })();
  return panel;
}

function annotateDisappearance(
  versions: readonly SavedLineageVersion[], entries: HistoryEntry[], revisionRead: ReadonlySet<number>,
): void {
  const byVersion = versions.map((): HistoryEntry[] => []);
  for (const entry of entries) byVersion[entry.version].push(entry);
  const tracked = (entry: HistoryEntry) => entry.revision && [entry.kind, entry.revision.wordId,
    entry.date, entry.author, entry.text].join("\0");
  for (let index = 1; index < versions.length; index++) {
    const before = versions[index - 1], after = versions[index];
    if (before.unavailable || after.unavailable || after.gapBefore) continue;
    const earlier = byVersion[index - 1], later = byVersion[index];
    if (!before.commentsUnavailable && !after.commentsUnavailable) {
      const keys = new Set(later.flatMap(entry => entry.thread ? [entry.thread.key] : []));
      const beforeCounts = new Map<string, number>();
      for (const thread of before.threads) beforeCounts.set(thread.key, (beforeCounts.get(thread.key) ?? 0) + 1);
      for (const entry of earlier) {
        if (entry.thread?.key.startsWith("d:") && beforeCounts.get(entry.thread.key) === 1 &&
            !keys.has(entry.thread.key)) entry.absentAt = index;
      }
    }
    if (!revisionRead.has(index) || !revisionRead.has(index - 1)) continue;
    const keys = new Set(later.map(tracked));
    const counts = new Map<string, number>();
    for (const entry of earlier) {
      const key = tracked(entry);
      if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    const currentIds = new Set(later.flatMap(entry => entry.revision?.wordId ? [entry.revision.wordId] : []));
    const currentText = new Set(later.flatMap(entry => entry.revision ? [entry.text] : []));
    for (const entry of earlier) {
      const revision = entry.revision;
      const key = tracked(entry);
      if (!key || !revision?.wordId || !entry.ranges.length || counts.get(key) !== 1 ||
          keys.has(key) || currentIds.has(revision.wordId) || currentText.has(entry.text)) continue;
      const aligned = entry.ranges.every(range => {
        const source = before.paragraphs.find(p => p.index === range.paragraph);
        return source?.paraId && after.paragraphs.filter(p => p.paraId === source.paraId && p.text === source.text).length === 1;
      });
      if (aligned) entry.absentAt = index;
    }
  }
}
