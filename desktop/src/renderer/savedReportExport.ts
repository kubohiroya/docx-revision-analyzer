import { DEFAULT_REPORT_RANGE, REPORT_LIMITS, type SavedVersionReport } from "../reportData";
import { SavedReportAnalysis, type SavedReportScope } from "../savedReport";
import type { AppLang, SavedLineageVersion } from "../shared";
import { compareSavedVersionPath, inferSavedVersionTimeline } from "../versionReviewPath";
import { el, modal } from "./extensions";
import { savedReportHtml } from "./savedReportHtml";
import { loadSavedVersionChart, savedReportFigures, type FigureKind, type FigureSettings } from "./savedReportFigures";
import { savedReportTimeline } from "./savedReportTimeline";
import { renderSavedVersionReport, type ReportViewState, type ReviewCommentDraft } from "./savedReportView";

const LABELS = {
  en: {
    review: "Review changes and comments", preview: "HTML preview", back: "Review and comments", save: "Save HTML…", close: "Close",
    saving: "Saving…", saved: "HTML saved", failed: "Could not prepare the report", saveFailed: "Could not save the report",
    openFailed: "Could not open the original file", exportCount: (versions: number, changes: number) => `Export · ${versions} versions / ${changes} changes`,
    previewHint: "This is what will be saved. Open Text range to inspect the full exported excerpt.",
    commentSaved: (backup: string) => `Word comment saved. Copy from before writing: ${backup}`,
    path: "Path to compare", allPaths: "All versions in scope (by date)",
    interval: "Interval", allIntervals: "Entire path",
    noPaths: "No branch paths could be inferred from the selected versions.",
    pathUnavailable: "The selected path is unavailable. Choose another path to compare.",
  },
  ja: {
    review: "変更箇所の確認とコメント", preview: "HTMLプレビュー", back: "確認・コメント", save: "HTMLを保存…", close: "閉じる",
    saving: "保存しています…", saved: "HTMLを保存しました", failed: "レポートの準備に失敗しました", saveFailed: "レポートの保存に失敗しました",
    openFailed: "元のファイルを開けませんでした", exportCount: (versions: number, changes: number) => `書き出す内容 · ${versions}版 / ${changes}箇所`,
    previewHint: "保存する内容のプレビューです。「表示範囲」を開くと、収録する原文の全範囲を確認できます。",
    commentSaved: (backup: string) => `Wordコメントを追記しました。追記前のバックアップ: ${backup}`,
    path: "比較する経路", allPaths: "対象とする版すべて（時刻順）",
    interval: "比較する区間", allIntervals: "経路全体",
    noPaths: "対象とする版から経路を推定できませんでした。",
    pathUnavailable: "選択した経路を利用できません。比較する経路を選び直してください。",
  },
};

export function savedReportExportLabel(lang: AppLang): string { return LABELS[lang].review; }

export async function showSavedReportExport(
  analysis: SavedReportAnalysis, versions: readonly SavedLineageVersion[], labels: string[],
  lang: AppLang, figureSettings: FigureSettings, branchTimelineEnabled: boolean,
  onVersionUpdated?: (index: number, version: SavedLineageVersion) => void,
): Promise<void> {
  const s = LABELS[lang];
  const canWrite = await window.app.reviewCommentWriteEnabled?.().catch(() => false) === true;
  const currentVersions = [...versions];
  const sameNames = new Set<string>();
  const seenNames = new Set<string>();
  for (const { source, label } of versions) {
    if (source !== "selected") continue;
    if (seenNames.has(label)) sameNames.add(label);
    seenNames.add(label);
  }
  const authors = (indices: readonly number[]) => [...new Set(indices.flatMap(index =>
    currentVersions[index].threads.flatMap(thread => thread.detail.comments
      .flatMap(comment => comment.author?.trim() ? [comment.author.trim()] : []))))].sort(new Intl.Collator(lang).compare);
  type IncludedChanges = Map<string, [number, number]>;
  const pathSelections = new Map<string, IncludedChanges>();
  let included: IncludedChanges = new Map();
  pathSelections.set("", included);
  let requestedPath = "";
  let intervalEnd: number | undefined;
  let paths = new Map<string, number[]>();
  let pathIndices: readonly number[] | undefined;
  let pathAnalysis: SavedReportAnalysis | undefined;
  const reviewState: ReportViewState = {};
  const commentDrafts = new Map<string, ReviewCommentDraft>();
  let targets = currentVersions.map((_, index) => index);
  const activeAnalysis = () => pathAnalysis ?? analysis;
  const activeTargets = () => pathIndices ? pathIndices.map((_, index) => index) : targets;
  const originalIndex = (index: number) => pathIndices ? pathIndices[index] : index;
  const analysisIndex = (index: number) => pathIndices ? pathIndices.indexOf(index) : index;
  let outputRange = DEFAULT_REPORT_RANGE;
  let preview: SavedVersionReport | undefined;
  let saving = false;
  let writing = false;
  let thumbnailView: AbortController | undefined;
  const thumbnails = new Map<string, Promise<string | undefined>>();
  function loadChart(index: number) {
    const version = currentVersions[index];
    const id = version.openId;
    if (!id) return Promise.resolve(undefined);
    let graphic = thumbnails.get(id);
    if (!graphic) {
      graphic = loadSavedVersionChart(version, figureSettings).catch(() => undefined);
      thumbnails.set(id, graphic);
    }
    return graphic;
  }

  return modal<void>((dialog, close) => {
    const wrapper = el("div", undefined, "report-export");
    const actions = el("div", undefined, "report-export-actions");
    const status = el("p", undefined, "hint");
    status.setAttribute("role", "status");
    const count = el("span", undefined, "hint report-export-count");
    count.setAttribute("aria-live", "polite");
    const content = el("div");
    const pathBox = el("div", undefined, "report-path-controls");
    const pathSelect = el("select");
    const pathLabel = el("label", s.path);
    pathLabel.append(pathSelect);
    const intervalSelect = el("select");
    const intervalLabel = el("label", s.interval);
    intervalLabel.append(intervalSelect);
    const controls = el("div", undefined, "report-branch-selectors");
    controls.append(pathLabel, intervalLabel);
    const pathDescription = el("p", undefined, "hint");
    pathDescription.setAttribute("role", "status");
    const versionName = (index: number) => {
      const { source, label } = currentVersions[index];
      if (source === "archive") return labels[index];
      return sameNames.has(label) ? `${index + 1}. ${label}` : label;
    };
    const timeline = savedReportTimeline(lang, currentVersions, versionName, choosePath, (after) => {
      const end = analysisIndex(after);
      if (end > 0) chooseInterval(intervalEnd === end ? undefined : end);
    });
    pathBox.append(timeline.element, controls, pathDescription);
    pathSelect.onchange = () => choosePath(pathSelect.value);
    intervalSelect.onchange = () => chooseInterval(intervalSelect.value ? Number(intervalSelect.value) : undefined);

    function chooseInterval(end: number | undefined): void {
      if (intervalEnd === end) return;
      intervalEnd = end;
      intervalSelect.value = end === undefined ? "" : String(end);
      reviewState.selected = undefined;
      showReview();
    }

    function choosePath(key: string): void {
      if (key === requestedPath) return;
      try {
        applyPath(key);
        reviewState.selected = undefined;
        showReview();
      } catch (error) {
        pathSelect.value = requestedPath;
        status.textContent = `${s.failed} · ${String(error)}`;
      }
    }
    const hint = el("p", s.previewHint, "hint");
    const action = (label: string, click: () => void, css?: string) => {
      const node = el("button", label, css);
      node.type = "button";
      node.onclick = click;
      return node;
    };
    const reviewButton = action(s.back, () => { if (preview && !saving) showReview(); });
    const previewButton = action(s.preview, showPreview);
    const save = action(s.save, () => void savePreview(), "primary");
    const closeButton = (click: () => void) => {
      const label = el("span", "×");
      label.setAttribute("aria-hidden", "true");
      const button = action("", click, "report-close-icon");
      button.setAttribute("aria-label", s.close);
      button.title = s.close;
      button.append(label);
      return button;
    };
    actions.append(el("h2", s.review), reviewButton, previewButton, save, closeButton(() => close()));
    wrapper.append(actions, count, pathBox, hint, status, content);
    dialog.append(wrapper);

    function showFigures(trigger: HTMLButtonElement, index: number, kind: FigureKind): void {
      if (writing || saving || preview) return;
      const indices = targets.includes(index) ? [...targets].sort((a, b) => a - b) : [index];
      const scroll = dialog.scrollTop;
      const dismiss = () => {
        panel.remove();
        dialog.removeEventListener("cancel", onCancel, true);
        wrapper.hidden = false;
        dialog.scrollTop = scroll;
        trigger.focus({ preventScroll: true });
      };
      const onCancel = (event: Event) => {
        event.preventDefault();
        event.stopImmediatePropagation();
        dismiss();
      };
      const panel = savedReportFigures(lang, indices.map(index => currentVersions[index]), indices.indexOf(index), kind,
        figureSettings, closeButton(dismiss), dismiss);
      wrapper.hidden = true;
      dialog.append(panel);
      dialog.addEventListener("cancel", onCancel, true);
      dialog.scrollTop = 0;
      panel.querySelector("select")?.focus({ preventScroll: true });
    }

    function applyPath(key: string, force = false): void {
      const next = key ? paths.get(key) ?? [] : undefined;
      if (!force && key === requestedPath && next?.join(":") === pathIndices?.join(":")) return;
      const nextAnalysis = next && compareSavedVersionPath(currentVersions, next, analysis.level);
      if (key !== requestedPath) intervalEnd = undefined;
      requestedPath = key;
      pathIndices = next;
      pathAnalysis = nextAnalysis;
      included = pathSelections.get(key) ?? new Map();
      pathSelections.set(key, included);
      const path = next ?? [];
      intervalSelect.replaceChildren(option("", s.allIntervals), ...path.slice(1).map((index, before) =>
        option(String(before + 1), `${versionName(path[before])} → ${versionName(index)}`)));
      intervalSelect.value = intervalEnd === undefined ? "" : String(intervalEnd);
    }

    function option(value: string, label: string) {
      const node = el("option", label);
      node.value = value;
      return node;
    }

    function refreshPaths(): void {
      if (!branchTimelineEnabled) return;
      const graph = inferSavedVersionTimeline(currentVersions, targets);
      paths = graph.paths;
      applyPath(requestedPath);
      timeline.refresh(graph.entries, paths);
      const choices = [...paths].map(([key, path]) => option(key, `${versionName(path[0])} → ${versionName(path[path.length - 1])}`));
      if (requestedPath && !paths.has(requestedPath)) choices.push(option(requestedPath, s.pathUnavailable));
      pathSelect.replaceChildren(option("", s.allPaths), ...choices);
      pathSelect.value = requestedPath;
      updatePathDescription();
    }

    function updatePathDescription(): void {
      pathDescription.textContent = pathIndices?.length ? pathIndices.map(versionName).join(" → ")
        : requestedPath ? s.pathUnavailable : paths.size ? "" : s.noPaths;
      pathDescription.hidden = !pathDescription.textContent;
    }

    function endpoints(indices: readonly number[]): number[] {
      return indices.length >= 2 ? [indices[0], indices[indices.length - 1]] : [...indices];
    }

    function timelinePairs(): [number, number][] {
      const interval = activeInterval();
      if (interval) return [interval];
      const indices = activeTargets();
      return indices.slice(1).map((after, index) => [indices[index], after]);
    }

    function activeInterval(): [number, number] | undefined {
      if (intervalEnd && pathIndices && intervalEnd < pathIndices.length) return [intervalEnd - 1, intervalEnd];
    }

    function exportSelection() {
      const valid = new Set(timelinePairs().map(pair => pair.join(":")));
      const selected = [...included].filter(([, pair]) => valid.has(pair.join(":")));
      return new Map(selected);
    }

    function scope(review: boolean): SavedReportScope {
      const selection = review ? undefined : exportSelection();
      const indices = activeTargets();
      const includePath = review || (pathIndices !== undefined && !selection?.size);
      const versions = activeInterval() ?? (includePath ? indices : selection?.size
        ? [...new Set([...selection.values()].flat())].sort((a, b) => a - b) : endpoints(indices));
      const selectedEdges = new Set([...(selection?.values() ?? [])].map(pair => pair.join(":")));
      const onlyPairs = includePath ? timelinePairs() : timelinePairs().filter(pair => selectedEdges.has(pair.join(":")));
      if (!review && !pathIndices && !onlyPairs.length && versions.length === 2) onlyPairs.push([versions[0], versions[1]]);
      const excluded = new Set<string>();
      if (selection) {
        for (const [before, after] of onlyPairs) {
          for (const change of activeAnalysis().changes(before, after)) if (!selection.has(change.id)) excluded.add(change.id);
        }
      }
      return { versions, initial: [versions[0], versions[versions.length - 1]], onlyPairs, excluded, includeText: true,
        range: review ? REPORT_LIMITS.length - 1 : outputRange, includeComments: true,
        includeDeleted: true, includeLinks: true, includeUnlocated: review, includeRelatedSources: true, review,
        observationTimeline: onlyPairs.length > 1 };
    }

    function updateMode(): void {
      reviewButton.setAttribute("aria-pressed", String(!preview));
      previewButton.setAttribute("aria-pressed", String(Boolean(preview)));
      const indices = activeTargets();
      previewButton.disabled = indices.length < 2 || saving || writing;
      pathBox.hidden = !branchTimelineEnabled || Boolean(preview);
      pathSelect.value = requestedPath;
      pathSelect.disabled = saving || writing || (!paths.size && !requestedPath);
      intervalLabel.hidden = !pathIndices?.length;
      intervalSelect.disabled = saving || writing;
      if (branchTimelineEnabled) timeline.setDisabled(saving || writing);
      reviewButton.disabled = saving;
      save.hidden = hint.hidden = !preview;
      save.disabled = saving;
      const selection = exportSelection();
      const countVersions = selection.size ? new Set([...selection.values()].flat()).size
        : activeInterval() ? 2 : pathIndices ? indices.length : endpoints(indices).length;
      count.textContent = s.exportCount(countVersions, selection.size);
    }

    function activeLabels(): readonly string[] {
      return pathIndices?.map(index => labels[index]) ?? labels;
    }

    function projectReview(selected: readonly number[]) {
      if (selected !== targets) {
        targets = [...selected];
        refreshPaths();
      }
      const configuration = scope(true);
      const report = activeAnalysis().project(configuration, lang, activeLabels());
      for (const thread of report.threads) {
        if (thread.reviewVersion !== undefined) thread.reviewVersion = originalIndex(thread.reviewVersion);
      }
      const sourceIndices = configuration.versions.map(originalIndex);
      updatePathDescription();
      if (branchTimelineEnabled) {
        const interval = activeInterval();
        timeline.select(pathIndices, interval && [originalIndex(interval[0]), originalIndex(interval[1])]);
      }
      updateMode();
      return { report, sourceIndices, authors: authors(pathIndices ?? currentVersions.map((_, index) => index)) };
    }

    function openLink(url: string): void {
      void window.app.watch.openCommentLink(url).catch((error) => { status.textContent = String(error); });
    }

    function showReview(): void {
      thumbnailView?.abort();
      thumbnailView = new AbortController();
      preview = undefined;
      status.textContent = "";
      try {
        const projection = projectReview(targets);
        renderSavedVersionReport(content, projection.report, {
          openLink, state: reviewState,
          review: {
            drafts: analysis.drafts(lang, labels), targets, projection, project: projectReview, outputRange,
            included: (id) => included.has(id),
            setIncluded: (id, include, before, after) => {
              if (include) included.set(id, [analysisIndex(before), analysisIndex(after)]);
              else included.delete(id);
              updateMode();
            },
            setOutputRange: (range) => { outputRange = range; },
            figures: { loadChart, open: showFigures, signal: thumbnailView.signal },
            canOpenVersion: (index) => Boolean(currentVersions[index]?.openId),
            openVersion: (index) => {
              const id = currentVersions[index]?.openId;
              if (id) void window.app.openSavedVersion(id).catch((error) => {
                status.textContent = `${s.openFailed} · ${String(error)}`;
              });
            },
            commentComposer: canWrite ? commentComposer : undefined,
          },
        });
      } catch (error) {
        status.textContent = `${s.failed} · ${String(error)}`;
      }
    }

    function commentComposer(id: string, before: number, after: number) {
      const source = currentVersions[after];
      if (source?.source !== "selected" || !source.openId || source.commentsUnavailable) return;
      const target = activeAnalysis().commentTarget(analysisIndex(before), analysisIndex(after), id);
      if (!target) return;
      const key = `${requestedPath}|${id}`;
      let draft = commentDrafts.get(key);
      if (!draft) {
        draft = { text: "" };
        commentDrafts.set(key, draft);
      }
      const local = { id: key, index: after, openId: source.openId, target, draft };
      return { filename: source.label, paragraph: target.paragraph, draft, submit: () => void saveComment(local) };
    }

    async function saveComment({ id, index, openId, target, draft }: {
      id: string;
      index: number;
      openId: string;
      target: NonNullable<ReturnType<SavedReportAnalysis["commentTarget"]>>;
      draft: ReviewCommentDraft;
    }): Promise<void> {
      const author = draft.author?.trim();
      if (writing || preview || !author || !draft.text.trim()) return;
      writing = draft.pending = true;
      draft.failure = undefined;
      status.textContent = "";
      updateMode();
      let backup: string | undefined;
      try {
        const result = await window.app.addSavedVersionComment(openId, { author, text: draft.text, range: target.range });
        if (!result.ok) {
          draft.failure = result.reason;
          return;
        }
        backup = result.backupPath;
        const { label, sourcePathHash } = currentVersions[index];
        const version = { ...result.version, label, sourcePathHash };
        currentVersions[index] = version;
        analysis = analysis.withVersions(currentVersions);
        applyPath(requestedPath, true);
        reviewState.author = author;
        commentDrafts.delete(id);
        onVersionUpdated?.(index, version);
      } catch {
        if (!backup) draft.failure = "saveFailed";
      } finally {
        writing = draft.pending = false;
        if (wrapper.isConnected && dialog.open) {
          showReview();
          if (backup) status.textContent = s.commentSaved(backup);
        }
      }
    }

    function showPreview(): void {
      if (activeTargets().length < 2 || saving || writing || preview) return;
      status.textContent = "";
      try {
        const configuration = scope(false);
        const candidate = activeAnalysis().project(configuration, lang, activeLabels());
        const selection = exportSelection();
        const index = configuration.onlyPairs?.flatMap(([before, after]) => activeAnalysis().changes(before, after)
          .filter((change) => selection.has(change.id))).findIndex((change) => change.id === reviewState.selected) ?? -1;
        renderSavedVersionReport(content, candidate, { openLink, initialChange: Math.max(0, index), state: {
          range: reviewState.range, commentsOpen: reviewState.commentsOpen,
          expanded: reviewState.expanded && [...reviewState.expanded],
          accordionHeights: reviewState.accordionHeights && [...reviewState.accordionHeights],
          table: reviewState.table && { ...reviewState.table, widths: [...reviewState.table.widths] },
        } });
        thumbnailView?.abort();
        preview = candidate;
        updateMode();
        save.focus();
      } catch (error) {
        status.textContent = `${s.failed} · ${String(error)}`;
      }
    }

    async function savePreview(): Promise<void> {
      if (!preview || saving) return;
      saving = true;
      updateMode();
      status.textContent = s.saving;
      try {
        const file = await window.app.saveFile("saved-versions.html", savedReportHtml(preview), "html");
        status.textContent = file ? `${s.saved} · ${file}` : "";
      } catch (error) {
        status.textContent = `${s.saveFailed} · ${String(error)}`;
      } finally {
        saving = false;
        updateMode();
      }
    }

    refreshPaths();
    showReview();
  }, undefined).finally(() => thumbnailView?.abort());
}
