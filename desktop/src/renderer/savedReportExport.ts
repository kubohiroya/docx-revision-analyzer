import { DEFAULT_REPORT_RANGE, REPORT_LIMITS, type SavedVersionReport } from "../reportData";
import { SavedReportAnalysis, type SavedReportScope } from "../savedReport";
import type { AppLang, SavedLineageVersion } from "../shared";
import { el, modal } from "./extensions";
import { savedReportHtml } from "./savedReportHtml";
import { loadSavedVersionChart, savedReportFigures, type FigureKind, type FigureSettings } from "./savedReportFigures";
import { renderSavedVersionReport, type ReportViewState, type ReviewCommentDraft } from "./savedReportView";

const LABELS = {
  en: {
    review: "Review changes and comments", preview: "HTML preview", back: "Review and comments", save: "Save HTML…", close: "Close",
    saving: "Saving…", saved: "HTML saved", failed: "Could not prepare the report", saveFailed: "Could not save the report",
    openFailed: "Could not open the original file", exportCount: (versions: number, changes: number) => `Export · ${versions} versions / ${changes} changes`,
    previewHint: "This is what will be saved. Open Text range to inspect the full exported excerpt.",
    commentSaved: (backup: string) => `Word comment saved. Copy from before writing: ${backup}`,
  },
  ja: {
    review: "変更箇所の確認とコメント", preview: "HTMLプレビュー", back: "確認・コメント", save: "HTMLを保存…", close: "閉じる",
    saving: "保存しています…", saved: "HTMLを保存しました", failed: "レポートの準備に失敗しました", saveFailed: "レポートの保存に失敗しました",
    openFailed: "元のファイルを開けませんでした", exportCount: (versions: number, changes: number) => `書き出す内容 · ${versions}版 / ${changes}箇所`,
    previewHint: "保存する内容のプレビューです。「表示範囲」を開くと、収録する原文の全範囲を確認できます。",
    commentSaved: (backup: string) => `Wordコメントを追記しました。追記前のバックアップ: ${backup}`,
  },
};

export function savedReportExportLabel(lang: AppLang): string { return LABELS[lang].review; }

export async function showSavedReportExport(
  analysis: SavedReportAnalysis, versions: readonly SavedLineageVersion[], labels: string[],
  lang: AppLang, figureSettings: FigureSettings, onVersionUpdated?: (index: number, version: SavedLineageVersion) => void,
): Promise<void> {
  const s = LABELS[lang];
  const canWrite = await window.app.reviewCommentWriteEnabled?.().catch(() => false) === true;
  const currentVersions = [...versions];
  const authors = () => [...new Set(currentVersions.flatMap((version) => version.threads.flatMap((thread) => thread.detail.comments
    .flatMap((comment) => comment.author?.trim() ? [comment.author.trim()] : []))))].sort(new Intl.Collator(lang).compare);
  const included = new Map<string, [number, number]>();
  const reviewState: ReportViewState = {};
  const commentDrafts = new Map<string, ReviewCommentDraft>();
  let targets = currentVersions.map((_, index) => index);
  let picked = targets.length >= 2 ? [0, targets.length - 1] : [...targets];
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
    wrapper.append(actions, count, hint, status, content);
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

    function timelinePairs(): [number, number][] {
      return targets.slice(1).map((after, index) => [targets[index], after]);
    }

    function exportSelection() {
      const valid = new Set(timelinePairs().map(pair => pair.join(":")));
      const selected = [...included].filter(([, pair]) => valid.has(pair.join(":")));
      return new Map(selected);
    }

    function scope(review: boolean): SavedReportScope {
      const selection = review ? undefined : exportSelection();
      const versions = review ? targets : selection?.size
        ? [...new Set([...selection.values()].flat())].sort((a, b) => a - b) : picked;
      const selectedEdges = new Set([...(selection?.values() ?? [])].map(pair => pair.join(":")));
      const onlyPairs = timelinePairs().filter(pair => review || selectedEdges.has(pair.join(":")));
      if (!review && !onlyPairs.length && picked.length === 2) onlyPairs.push([picked[0], picked[1]]);
      const excluded = new Set<string>();
      if (selection) {
        for (const [before, after] of onlyPairs) {
          for (const change of analysis.changes(before, after)) if (!selection.has(change.id)) excluded.add(change.id);
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
      previewButton.disabled = picked.length !== 2 || saving || writing;
      reviewButton.disabled = saving;
      save.hidden = hint.hidden = !preview;
      save.disabled = saving;
      const selection = exportSelection();
      const countVersions = selection.size ? new Set([...selection.values()].flat()).size : picked.length;
      count.textContent = s.exportCount(countVersions, selection.size);
    }

    function projectReview(versions: number[]): SavedVersionReport {
      picked = versions;
      const report = analysis.project(scope(true), lang, labels);
      updateMode();
      return report;
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
        renderSavedVersionReport(content, projectReview(picked), {
          openLink, state: reviewState,
          review: {
            drafts: analysis.drafts(lang, labels), targets, picked, project: projectReview,
            setTargets: (selected) => { targets = selected; }, authors: authors(), outputRange,
            included: (id) => included.has(id),
            setIncluded: (id, include, before, after) => {
              if (include) included.set(id, [before, after]);
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
      const target = analysis.commentTarget(before, after, id);
      if (!target) return;
      let draft = commentDrafts.get(id);
      if (!draft) {
        draft = { text: "" };
        commentDrafts.set(id, draft);
      }
      const local = { id, index: after, openId: source.openId, target, draft };
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
        const version = { ...result.version, label: currentVersions[index].label };
        currentVersions[index] = version;
        analysis = analysis.withVersions(currentVersions);
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
      if (picked.length !== 2 || saving || writing || preview) return;
      status.textContent = "";
      try {
        const configuration = scope(false);
        const candidate = analysis.project(configuration, lang, labels);
        const selection = exportSelection();
        const index = configuration.onlyPairs?.flatMap(([before, after]) => analysis.changes(before, after)
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

    showReview();
  }, undefined).finally(() => thumbnailView?.abort());
}
