import { DEFAULT_REPORT_RANGE, REPORT_LIMITS, type SavedVersionReport } from "../reportData";
import { SavedReportAnalysis, type SavedReportScope } from "../savedReport";
import type { AppLang, SavedLineageVersion } from "../shared";
import { el, modal } from "./extensions";
import { savedReportHtml } from "./savedReportHtml";
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
  initial: [number, number], lang: AppLang, onVersionUpdated?: (index: number, version: SavedLineageVersion) => void,
): Promise<void> {
  const s = LABELS[lang];
  const canWrite = await window.app.reviewCommentWriteEnabled?.().catch(() => false) === true;
  const currentVersions = [...versions];
  const authors = () => [...new Set(currentVersions.flatMap((version) => version.threads.flatMap((thread) => thread.detail.comments
    .flatMap((comment) => comment.author?.trim() ? [comment.author.trim()] : []))))].sort(new Intl.Collator(lang).compare);
  const included = new Set<string>();
  const reviewState: ReportViewState = {};
  const commentDrafts = new Map<string, ReviewCommentDraft>();
  let picked: number[] = [...initial];
  let outputRange = DEFAULT_REPORT_RANGE;
  let preview: SavedVersionReport | undefined;
  let saving = false;
  let writing = false;

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
    actions.append(el("h2", s.review), reviewButton, previewButton, save, action(s.close, () => close()));
    wrapper.append(actions, count, hint, status, content);
    dialog.append(wrapper);

    function scope(versions: number[], range: number, review: boolean): SavedReportScope {
      const [before, after] = [...versions].sort((a, b) => a - b);
      const excluded = review || after === undefined ? new Set<string>() : new Set(
        analysis.changes(before, after).map((change) => change.id).filter((id) => !included.has(id)),
      );
      return { versions, initial: [before, after], excluded, includeText: true, range, includeComments: true,
        includeDeleted: true, includeLinks: true, includeUnlocated: review, includeRelatedSources: true, review };
    }

    function updateMode(): void {
      reviewButton.setAttribute("aria-pressed", String(!preview));
      previewButton.setAttribute("aria-pressed", String(Boolean(preview)));
      previewButton.disabled = picked.length !== 2 || saving || writing;
      reviewButton.disabled = saving;
      save.hidden = hint.hidden = !preview;
      save.disabled = saving;
      const [before, after] = [...picked].sort((a, b) => a - b);
      const selected = after === undefined ? 0 : analysis.changes(before, after).filter((change) => included.has(change.id)).length;
      count.textContent = s.exportCount(picked.length, selected);
    }

    function projectReview(versions: number[]): SavedVersionReport {
      const report = analysis.project(scope(versions, REPORT_LIMITS.length - 1, true), lang, labels);
      picked = versions;
      updateMode();
      return report;
    }

    function openLink(url: string): void {
      void window.app.watch.openCommentLink(url).catch((error) => { status.textContent = String(error); });
    }

    function showReview(): void {
      preview = undefined;
      status.textContent = "";
      try {
        renderSavedVersionReport(content, projectReview(picked), {
          openLink, state: reviewState,
          review: { drafts: analysis.drafts(lang, labels), picked, project: projectReview, authors: authors(), outputRange,
            included: (id) => included.has(id),
            setIncluded: (id, include) => {
              if (include) included.add(id);
              else included.delete(id);
              updateMode();
            },
            setOutputRange: (range) => { outputRange = range; },
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

    function commentComposer(id: string) {
      const [before, after] = [...picked].sort((a, b) => a - b);
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
        const candidate = analysis.project(scope(picked, outputRange, false), lang, labels);
        const [before, after] = [...picked].sort((a, b) => a - b);
        const index = analysis.changes(before, after).filter((change) => included.has(change.id))
          .findIndex((change) => change.id === reviewState.selected);
        renderSavedVersionReport(content, candidate, { openLink, initialChange: Math.max(0, index), state: {
          range: reviewState.range, commentsOpen: reviewState.commentsOpen,
          expanded: reviewState.expanded && [...reviewState.expanded],
          table: reviewState.table && { ...reviewState.table, widths: [...reviewState.table.widths] },
        } });
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
  }, undefined);
}
