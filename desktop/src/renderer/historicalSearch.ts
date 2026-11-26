import { commentReferences } from "../../../src/lib/commentReferences";
import { projectHistoricalAnchor, type HistoricalMatch } from "../../../src/lib/historicalAnchors";
import { compareSavedVersions, type LineageRange, type SavedParagraphRange } from "../../../src/lib/lineage";
import { groupedHistoryRows, historyEntries, historyMatch, reconcileHistoryVersions, searchKey, type HistoryEntry } from "../historySearch";
import type { AppLang, SavedLineageVersion } from "../shared";
import { el } from "./extensions";

const labels = {
  en: {
    history: "Saved history", title: "Search saved history", search: "Text or quotation", all: "All", kind: "Kind", author: "Author", tag: "Tag", dateBasis: "Filter by date",
    version: "Version", from: "From", to: "Through", state: "State", resolved: "Resolved",
    absent: "No longer observed", unknownDate: "Unknown dates", open: "Show source", file: "Open source file",
    target: "Find corresponding position in", choose: "Choose a version", previous: "Previous", next: "Next",
    quote: "Recorded quotation", noPosition: "The original range could not be verified.", matched: "matches",
    observations: (count: number) => `${count} matching versions of this thread`,
    recorded: "Partial or recorded comments only", revisionCoverage: "Change history read", commentCoverage: "Comment files read", unavailable: "Unavailable", pending: "Pending", changed: "Changed since comparison", indexing: "Reading versions…",
    excluded: "Outside index scope", failed: "Unable to read change history", disabled: "Indexing disabled",
    tables: "tables", headersFooters: "headers and footers", notes: "footnotes and endnotes", textBoxes: "text boxes", formatting: "formatting changes",
    limited: "Positions and change history cover body paragraphs and headings. Text from unavailable versions may be missing.",
    notOrdered: "Choose a branch or compare versions in a confirmed order to find a corresponding position.",
    absentAt: "No longer observed in", range: "Corresponding range", partial: "Part of the original range",
    boundary: "Estimated boundary", related: "Related paragraph", ambiguous: "Multiple possible locations",
    reappeared: "Possible reappearance", reappearanceHint: "A passage with matching text reappears after an absence.",
    rangeHint: "Matching characters across the selected versions establish this span.",
    partialHint: "Only the highlighted characters could be followed across the selected versions.",
    boundaryHint: "Surviving text on both sides points to this boundary.",
    relatedHint: "The paragraph corresponds, but the original characters could not be matched.",
    ambiguousHint: "Multiple occurrences or a possible copy prevented an exact match.",
    noPath: "There was insufficient context to establish a match on this path.",
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
    observations: (count: number) => `このスレッドに一致した版 ${count} 件`,
    recorded: "一部または保全記録のコメントのみ", revisionCoverage: "変更履歴を読めた版", commentCoverage: "コメントを読めた版", unavailable: "読取不能", pending: "未処理", changed: "比較後に変更", indexing: "保存版を読み込んでいます…",
    excluded: "索引の対象外", failed: "変更履歴の読取失敗", disabled: "索引化が無効",
    tables: "表", headersFooters: "ヘッダー・フッター", notes: "脚注・文末脚注", textBoxes: "テキストボックス", formatting: "書式変更",
    limited: "位置照合と変更履歴の対象は本文と見出しです。読み取れない版の内容は検索に出ない場合があります。",
    notOrdered: "対応箇所を探すには、経路を選ぶか、順序を確認した版を比較してください。",
    absentAt: "観測されなくなった版", range: "対応する範囲", partial: "一部が対応する範囲",
    boundary: "推定した境界", related: "関連する段落", ambiguous: "複数の対応候補",
    reappeared: "再出現の候補", reappearanceHint: "いったん消えた箇所と同じ本文が後続版に現れています。",
    rangeHint: "選んだ版の間で一致した文字に基づく範囲です。",
    partialHint: "選んだ版の間でたどれた文字だけを強調しています。",
    boundaryHint: "境界の両側に残った本文から位置を推定しました。",
    relatedHint: "段落は対応していますが、元の文字の範囲は照合できませんでした。",
    ambiguousHint: "複数の出現や複製の可能性があるため、位置を絞れませんでした。",
    noPath: "この経路で対応を特定できる文脈が不足しています。",
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

type HistoryFilter = "query" | "kind" | "author" | "tag" | "version" | "state" | "dateBasis" | "from" | "to";

export interface HistorySearchState {
  filters?: Partial<Record<HistoryFilter, string>>;
  page?: number;
  selection?: string;
  target?: string;
  expanded?: Set<string>;
}

export function historySearchPanel(
  lang: AppLang, versions: readonly SavedLineageVersion[], names: readonly string[],
  ordered: boolean, anchors: boolean, search: boolean, signal: AbortSignal, dismiss: () => void, view: HistorySearchState = {},
): HTMLElement {
  const s = labels[lang];
  const panel = el("section", undefined, "report-export history-search");
  const header = el("div", undefined, "report-export-actions");
  header.append(el("h2", historySearchLabel(lang, search)), button(s.close, dismiss));
  const controls = el("div", undefined, "lineage-controls");
  const query = el("textarea");
  query.rows = 1;
  query.setAttribute("role", "searchbox");
  function field(label: string, control: HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement): void {
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
  const fields = Object.entries({ query, kind, author, tag, version, state, dateBasis, from, to }) as
    [HistoryFilter, HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement][];
  for (const [name, field] of fields) field.value = view.filters?.[name] ?? field.value;
  const expanded = view.expanded ??= new Set();
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
  const pageSize = 40;
  let displayed: HistoryEntry | undefined;
  let displayedAbsence: number | undefined;
  let pendingRender = false;
  let lineage: ReturnType<typeof compareSavedVersions> | undefined;
  const anchorVersions = [...versions];
  let shownTarget: number | undefined;
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

  function displayEntry(entry: HistoryEntry, restore = false): void {
    displayed = entry;
    displayedAbsence = entry.absentAt;
    shownTarget = undefined;
    if (!restore) view.target = undefined;
    view.selection = entry.id;
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
        select.value = view.target ?? "";
        select.onchange = () => {
          view.target = select.value;
          refreshProjection();
        };
      }
    }
    refreshProjection();
    if (!restore) detail.scrollIntoView({ block: "nearest" });
  }

  function refreshProjection(): void {
    if (!anchors || !ordered || !displayed) return;
    const target = view.target ? Number(view.target) : undefined;
    if (target === undefined || displayed.version >= target || target >= anchorVersions.length) {
      projected.replaceChildren();
      shownTarget = undefined;
    } else if (target >= processed) {
      projected.replaceChildren(el("p", s.indexing, "hint"));
      shownTarget = undefined;
    } else if (shownTarget !== target) {
      projected.replaceChildren();
      lineage ??= compareSavedVersions(anchorVersions, { japaneseWordHints: true });
      showMatch(projectHistoricalAnchor(lineage, displayed.version, target, displayed.ranges), target);
      shownTarget = target;
    }
  }

  function showCandidates(candidates: readonly LineageRange[]): void {
    const byVersion = new Map<number, SavedParagraphRange[]>();
    for (const candidate of candidates) {
      const ranges = byVersion.get(candidate.version) ?? [];
      ranges.push(candidate);
      byVersion.set(candidate.version, ranges);
    }
    for (const [version, ranges] of byVersion) {
      projected.append(el("h4", names[version]));
      addRanges(projected, version, ranges, false);
    }
  }

  function showMatch(match: HistoricalMatch, target: number): void {
    projected.append(el("h3", names[target]), el("p", s[match.kind]));
    switch (match.kind) {
      case "unknown":
        projected.append(el("p", match.reason === "gap" ? s.gap : match.reason === "no-source" ? s.noPosition : s.noPath, "hint"));
        break;
      case "ambiguous":
      case "reappeared":
        projected.append(el("p", match.kind === "reappeared" ? s.reappearanceHint : s.ambiguousHint, "hint"));
        showCandidates(match.ranges);
        break;
      case "deleted":
        break;
      default:
        projected.append(el("p", s[`${match.kind}Hint`], "hint"));
        addRanges(projected, target, match.ranges, match.kind !== "related");
        if (match.reappearance?.length) {
          projected.append(el("p", s.reappearanceHint, "hint"));
          showCandidates(match.reappearance);
        }
    }
  }

  function formatDate(iso: string): string {
    const time = Date.parse(iso);
    return Number.isFinite(time) ? new Date(time).toLocaleString(lang) : iso;
  }

  function render(): void {
    if (signal.aborted) return;
    const needle = searchKey(query.value);
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
          (needle && !entry.foldedText.includes(needle) && !entry.foldedQuote.includes(needle))) return false;
      const date = dateFor(entry);
      if (!date) unknown++;
      return (!from.value || (date !== "" && date >= from.value)) && (!to.value || (date !== "" && date <= to.value));
    });
    const rows = groupedHistoryRows(filtered);
    const page = Math.min(view.page ?? 0, Math.max(0, Math.ceil(rows.length / pageSize) - 1));
    status.textContent = [
      `${filtered.length} ${s.matched}`, `${s.revisionCoverage} ${read}/${versions.length}`,
      `${s.commentCoverage} ${commentsRead}/${versions.length}`, unknown ? `${s.unknownDate} ${unknown}` : "",
      processed < versions.length ? s.indexing : "",
    ].filter(Boolean).join(" · ");
    list.replaceChildren(...rows.slice(page * pageSize, (page + 1) * pageSize)
      .map(entries => renderRow(entries, needle)));
    const back = button(s.previous, () => { view.page = page - 1; render(); });
    back.disabled = page === 0;
    const next = button(s.next, () => { view.page = page + 1; render(); });
    next.disabled = (page + 1) * pageSize >= rows.length;
    paging.replaceChildren(back, next);

    const selected = view.selection && filtered.find(entry => entry.id === view.selection);
    if (selected && (selected !== displayed || selected.absentAt !== displayedAbsence)) displayEntry(selected, true);
    else if (displayed && shownTarget === undefined && view.target) refreshProjection();
  }

  function renderRow(entries: HistoryEntry[], needle: string): HTMLLIElement {
    if (entries.length === 1) return resultRow(entries[0], needle);
    const versionCount = new Set(entries.map(entry => entry.version)).size;
    const row = el("li");
    const group = el("details");
    const key = entries[0].threadGroup!;
    const author = entries[0].thread?.detail.comments[0]?.author;
    group.open = expanded.has(key) || entries.some(entry => entry.id === view.selection);
    const results = el("ul");
    function showObservations(limit: number): void {
      results.replaceChildren(...entries.slice(0, limit).map(entry => resultRow(entry, needle)));
      if (limit < entries.length) {
        const more = el("li");
        more.append(button(s.next, () => showObservations(limit + pageSize)));
        results.append(more);
      }
    }
    if (group.open) showObservations(pageSize);
    group.ontoggle = () => {
      if (signal.aborted || !group.isConnected) return;
      if (group.open) {
        expanded.add(key);
        if (!results.childElementCount) showObservations(pageSize);
      } else expanded.delete(key);
    };
    group.append(el("summary", [s.observations(versionCount), author].filter(Boolean).join(" · ")), results);
    row.append(group);
    return row;
  }

  function resultRow(entry: HistoryEntry, needle: string): HTMLLIElement {
    const row = el("li");
    const desc = [names[entry.version], s[entry.kind], entry.author, entry.resolved ? s.resolved : undefined,
      entry.absentAt !== undefined ? s.absent : undefined].filter(Boolean).join(" · ");
    const heading = button(desc, () => displayEntry(entry));
    heading.setAttribute("aria-label", `${s.open} · ${desc}`);
    row.append(heading);
    const match = historyMatch(entry.text, needle);
    const quotation = match ? undefined : historyMatch(entry.quote ?? "", needle);
    const hit = match ?? quotation;
    if (hit) {
      const text = match ? entry.text : entry.quote!;
      const start = Math.max(0, hit.start - 60);
      row.append(marked(text.slice(start, hit.end + 100), hit.start - start, hit.end - start));
    } else row.append(el("p", entry.text.slice(0, 160), "lineage-passage"));
    return row;
  }

  function scheduleRender(): void {
    if (pendingRender) return;
    pendingRender = true;
    requestAnimationFrame(() => {
      pendingRender = false;
      render();
    });
  }
  controls.hidden = !search;
  controls.oninput = () => {
    view.filters = Object.fromEntries(fields.map(([name, field]) => [name, field.value]));
    view.page = 0;
    view.selection = view.target = undefined;
    displayed = undefined;
    detail.replaceChildren(); choices.replaceChildren(); projected.replaceChildren();
    scheduleRender();
  };
  render();

  void (async () => {
    let previous: Parameters<typeof reconcileHistoryVersions>[0] | undefined;
    for (const [index, version] of versions.entries()) {
      if (signal.aborted) return;
      const result = version.openId ? await window.app.readSavedVersionRevisions(version.openId).catch(() => undefined) : undefined;
      if (signal.aborted) return;
      const changesRead = result?.status === "read";
      if (changesRead) read++;
      const verified = !version.openId || (result !== undefined && result.status !== "changed" && result.status !== "unavailable");
      if (version.openId && (!verified || result?.status === "failed")) {
        anchorVersions[index] = { ...version, unavailable: true };
        lineage = undefined;
      }
      const hasComments = verified && !version.unavailable && !version.commentsUnavailable;
      if (hasComments) commentsRead++;
      const entries = historyEntries(verified ? version : { ...version, threads: [] }, index, result?.records ?? []);
      if (result?.status === "failed") for (const entry of entries) entry.ranges = [];
      const revisionStatus = result?.status ?? "unavailable";
      const commentStatus = hasComments ? s.commentCoverage : verified && version.threads.length ? s.recorded : s.unavailable;
      coverage[index].textContent = [
        names[index],
        revisionStatus === "read" ? s.revisionCoverage : s[revisionStatus],
        result?.status === "read" && result.excluded.length
          ? `${s.excluded} (${result.excluded.map(region => s[region]).join(", ")})` : "",
        commentStatus,
      ].filter(Boolean).join(" · ");
      const current = { version, entries, changesRead, commentsRead: hasComments };
      if (ordered && previous) reconcileHistoryVersions(previous, current, index);
      previous = current;
      for (const entry of entries) history.push(entry);
      processed++;
      scheduleRender();
      // Yield to input and painting even when the next version has no file to read.
      await new Promise<void>(resolve => setTimeout(resolve, 0));
    }
  })();
  return panel;
}
