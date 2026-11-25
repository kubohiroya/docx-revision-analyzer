import { DEFAULT_REPORT_RANGE, type SavedVersionReport } from "../reportData";
import { SavedReportAnalysis, reportPairs, type SavedReportScope } from "../savedReport";
import type { AppLang, SavedLineageVersion } from "../shared";
import { el, modal } from "./extensions";
import { savedReportHtml } from "./savedReportHtml";
import { REPORT_LABELS, renderSavedVersionReport, reportChangeLabel } from "./savedReportView";

const EN = {
  export: "Export HTML", drafts: "Drafts to export", changes: "Changes to export", pair: "Changes between",
  include: "Include", filename: "Filename", createdAt: "Created", updatedAt: "Modified", archivedAt: "Archived",
  bytes: "Size (B)", characters: "Characters", words: "Words",
  countsHint: "Character and word counts use the visible body text and headings.",
  details: "Content to include in the HTML", text: "Headings and surrounding text", max: "Maximum text range to export",
  comments: "Word comments, replies and tags", deleted: "Include deleted threads", links: "Include comment reference URLs",
  unlocated: "Include comments without a confirmed location", pageNote: "A full page also includes the surrounding text on that page.",
  preview: "Preview HTML", back: "Back to export options", save: "Save HTML…", close: "Close",
  scopeHint: "Choose the content to include in the HTML file. Viewing filters are independent of the export options.",
  previewHint: "This is the HTML preview. Open Text range and move the slider to its highest position to inspect all the included text.",
  minimum: "Choose at least two drafts.", working: "Preparing the preview…", saving: "Saving…", saved: "HTML saved",
  failed: "Could not prepare the report", saveFailed: "Could not save the report", all: "Select all", none: "Deselect all",
  noChanges: "No changes are available for this pair.", selection: (n: number, total: number) => `${n} of ${total} selected`,
};

const LABELS: Record<AppLang, typeof EN> = { en: EN, ja: {
  export: "HTMLを書き出す", drafts: "書き出す稿", changes: "書き出す変更箇所", pair: "変更箇所を選ぶ稿の組",
  include: "出力", filename: "ファイル名", createdAt: "作成日時", updatedAt: "更新日時", archivedAt: "保管日時",
  bytes: "サイズ (B)", characters: "文字数", words: "単語数",
  countsHint: "文字数・単語数は、表示される本文と見出しをもとに数えます。",
  details: "HTMLに含める情報", text: "見出しと前後の原文", max: "書き出す原文の上限",
  comments: "Wordコメント・返信・タグ", deleted: "削除済みスレッドも含める", links: "コメントの参照URLを含める",
  unlocated: "位置を特定できないコメントも含める", pageNote: "1ページ全体を選ぶと、そのページの周辺の原文も含まれます。",
  preview: "HTMLをプレビュー", back: "書き出し範囲の選択に戻る", save: "HTMLを保存…", close: "閉じる",
  scopeHint: "比較画面の絞り込みとは別に、HTMLファイルに収録する内容を選びます。",
  previewHint: "保存するHTMLのプレビューです。「表示範囲」を開き、スライダーを右端にすると、収録する原文の全範囲を確認できます。",
  minimum: "書き出す稿を2つ以上選んでください。", working: "プレビューを準備しています…", saving: "保存しています…", saved: "HTMLを保存しました",
  failed: "レポートの準備に失敗しました", saveFailed: "レポートの保存に失敗しました", all: "すべて選択", none: "すべて解除",
  noChanges: "この2稿には書き出せる変更箇所がありません。", selection: (n, total) => `${total} 件中 ${n} 件を選択`,
} };

export function savedReportExportLabel(lang: AppLang): string { return LABELS[lang].export; }

export function showSavedReportExport(
  analysis: SavedReportAnalysis, versions: readonly SavedLineageVersion[], labels: string[],
  initial: [number, number], lang: AppLang,
): Promise<void> {
  const s = LABELS[lang];
  const viewLabels = REPORT_LABELS[lang];
  const excluded = new Set<string>();
  const scope: SavedReportScope = {
    versions: [...initial], initial, excluded, includeText: true, range: DEFAULT_REPORT_RANGE,
    includeComments: false, includeDeleted: false, includeLinks: false, includeUnlocated: false,
  };
  return modal<void>((dialog, close) => {
    const wrapper = el("div", undefined, "report-export");
    const form = el("section", undefined, "report-export-scope");
    const preview = el("section", undefined, "report-export-preview");
    const content = el("div");
    const notice = el("p", undefined, "hint");
    const feedback = el("p", undefined, "hint");
    notice.setAttribute("role", "status");
    feedback.setAttribute("role", "status");
    let report: SavedVersionReport | undefined;
    let pairChoice = initial.join(":");
    let busy = false;
    const action = (label: string, click: () => void, style?: string) => {
      const button = el("button", label, style);
      button.type = "button";
      button.onclick = click;
      return button;
    };
    const checkbox = (label: string, checked: boolean, change: (checked: boolean) => void) => {
      const input = el("input");
      input.type = "checkbox";
      input.checked = checked;
      input.onchange = () => change(input.checked);
      const field = el("label", undefined, "report-check");
      field.append(input, label);
      return { input, field };
    };
    const choice = (label: string, value: string) => {
      const item = el("option", label);
      item.value = value;
      return item;
    };

    const draftFields = el("fieldset");
    draftFields.append(el("legend", s.drafts));
    const tableWrap = el("div", undefined, "report-draft-table-wrap");
    const table = el("table", undefined, "report-draft-table");
    table.setAttribute("aria-label", s.drafts);
    const header = table.createTHead().insertRow();
    const hasArchive = versions.some((version) => version.archivedAt);
    for (const label of [s.include, s.filename, s.createdAt, s.updatedAt, ...(hasArchive ? [s.archivedAt] : []),
      s.bytes, s.characters, s.words]) {
      const column = el("th", label);
      column.scope = "col";
      header.append(column);
    }
    const numbers = new Intl.NumberFormat(lang);
    const dates = new Intl.DateTimeFormat(lang, { dateStyle: "short", timeStyle: "medium" });
    const segmenter = new Intl.Segmenter(lang, { granularity: "word" });
    const date = (iso?: string) => {
      const value = iso ? Date.parse(iso) : NaN;
      return Number.isFinite(value) ? dates.format(value) : "";
    };
    const body = table.createTBody();
    for (const [index, version] of versions.entries()) {
      const { input, field } = checkbox("", scope.versions.includes(index), (checked) => {
        scope.versions = checked ? [...scope.versions, index].sort((a, b) => a - b) : scope.versions.filter((v) => v !== index);
        updatePairs();
      });
      input.setAttribute("aria-label", `${s.include} · ${labels[index]}`);
      let characters = 0;
      let words = 0;
      if (!version.unavailable) {
        for (const paragraph of version.paragraphs) {
          characters += Array.from(paragraph.text).length;
          for (const segment of segmenter.segment(paragraph.text)) if (segment.isWordLike) words++;
        }
      }
      const row = body.insertRow();
      row.insertCell().append(field);
      const filename = el("th", version.label);
      filename.scope = "row";
      row.append(filename);
      for (const value of [date(version.fileCreatedAt), date(version.sourceModified),
        ...(hasArchive ? [date(version.archivedAt)] : []),
        version.fileSize === undefined ? "" : numbers.format(version.fileSize),
        version.unavailable ? "" : numbers.format(characters),
        version.unavailable ? "" : numbers.format(words)]) {
        row.insertCell().textContent = value;
      }
    }
    tableWrap.append(table);
    draftFields.append(tableWrap, el("p", s.countsHint, "hint"));

    const options = el("fieldset");
    options.append(el("legend", s.details));
    const text = checkbox(s.text, scope.includeText, (value) => { scope.includeText = value; updateOptions(); });
    const rangeLabel = el("label", s.max, "report-check report-export-indent");
    const range = el("select");
    viewLabels.ranges.forEach((label, index) => range.append(choice(label, String(index))));
    range.value = String(scope.range);
    range.onchange = () => { scope.range = Number(range.value); updateOptions(); };
    rangeLabel.append(range);
    const pageHint = el("p", s.pageNote, "hint report-export-indent");
    const comments = checkbox(s.comments, false, (value) => { scope.includeComments = value; updateOptions(); });
    const deleted = checkbox(s.deleted, false, (value) => { scope.includeDeleted = value; });
    const links = checkbox(s.links, false, (value) => { scope.includeLinks = value; });
    const unlocated = checkbox(s.unlocated, false, (value) => { scope.includeUnlocated = value; });
    for (const option of [deleted, links, unlocated]) option.field.classList.add("report-export-indent");
    options.append(text.field, rangeLabel, pageHint, comments.field, deleted.field, links.field, unlocated.field);

    const changes = el("fieldset");
    const selector = el("select");
    const selectorLabel = el("label", s.pair, "report-check");
    selectorLabel.append(selector);
    selector.onchange = () => { pairChoice = selector.value; updateChanges(); };
    const selectedCount = el("span", undefined, "hint");
    selectedCount.setAttribute("aria-live", "polite");
    const rows = el("div", undefined, "report-scope-changes");
    const changeControls = el("div", undefined, "report-export-actions");
    const markAll = (include: boolean) => {
      const [before, after] = pairChoice.split(":").map(Number);
      if (!pairChoice) return;
      for (const row of analysis.changes(before, after)) {
        if (include) excluded.delete(row.id);
        else excluded.add(row.id);
      }
      updateChanges();
    };
    const all = action(s.all, () => markAll(true));
    const none = action(s.none, () => markAll(false));
    changeControls.append(all, none, selectedCount);
    changes.append(el("legend", s.changes), selectorLabel, changeControls, rows);

    const formControls = el("div", undefined, "report-export-actions");
    const submit = action(s.preview, () => {
      if (busy || scope.versions.length < 2) return;
      busy = true;
      submit.disabled = true;
      notice.textContent = s.working;
      setTimeout(() => {
        if (!dialog.open) return;
        try {
          scope.initial = initial.every((v) => scope.versions.includes(v))
            ? initial : [scope.versions[0], scope.versions[1]];
          report = analysis.project(scope, lang, labels);
          renderSavedVersionReport(content, report, (url) => void window.app.watch.openCommentLink(url).catch((error) => {
            feedback.textContent = String(error);
          }));
          feedback.textContent = "";
          form.hidden = true;
          preview.hidden = false;
          dialog.scrollTop = 0;
          back.focus();
        } catch (error) {
          notice.textContent = `${s.failed} · ${String(error)}`;
        } finally {
          busy = false;
          submit.disabled = false;
        }
      }, 0);
    }, "primary");
    formControls.append(submit, action(s.close, () => close()));
    const grid = el("div", undefined, "report-export-options");
    grid.append(changes, options);
    form.append(el("h2", s.export), el("p", s.scopeHint, "hint"), draftFields, grid, notice, formControls);

    const previewControls = el("div", undefined, "report-export-actions");
    const back = action(s.back, () => {
      report = undefined;
      content.replaceChildren();
      preview.hidden = true;
      form.hidden = false;
      notice.textContent = "";
      dialog.scrollTop = 0;
      submit.focus();
    });
    const save = action(s.save, async () => {
      if (!report || busy) return;
      busy = true;
      save.disabled = true;
      back.disabled = true;
      feedback.textContent = s.saving;
      try {
        const file = await window.app.saveFile("saved-versions.html", savedReportHtml(report), "html");
        feedback.textContent = file ? `${s.saved} · ${file}` : "";
      } catch (error) {
        feedback.textContent = `${s.saveFailed} · ${String(error)}`;
      } finally {
        busy = false;
        save.disabled = false;
        back.disabled = false;
      }
    }, "primary");
    previewControls.append(back, save, action(s.close, () => close()));
    preview.append(previewControls, el("p", s.previewHint, "hint"), feedback, content);
    preview.hidden = true;
    wrapper.append(form, preview);
    dialog.append(wrapper);

    function updateOptions(): void {
      range.disabled = !scope.includeText;
      pageHint.hidden = !scope.includeText || scope.range !== viewLabels.ranges.length - 1;
      for (const option of [deleted, links, unlocated]) option.input.disabled = !scope.includeComments;
    }

    function updatePairs(): void {
      const pairs = reportPairs(scope.versions);
      const choices = pairs.map(([a, b]) => choice(`${labels[a]} → ${labels[b]}`, `${a}:${b}`));
      if (!choices.some((item) => item.value === pairChoice)) pairChoice = choices[0]?.value ?? "";
      selector.replaceChildren(...choices);
      selector.value = pairChoice;
      changes.hidden = !choices.length;
      submit.disabled = scope.versions.length < 2;
      notice.textContent = submit.disabled ? s.minimum : "";
      updateChanges();
    }

    function updateChanges(): void {
      rows.replaceChildren();
      const [before, after] = pairChoice.split(":").map(Number);
      const candidates = pairChoice ? analysis.changes(before, after) : [];
      const updateCount = () => {
        const count = candidates.reduce((n, change) => n + Number(!excluded.has(change.id)), 0);
        selectedCount.textContent = s.selection(count, candidates.length);
        all.disabled = count === candidates.length;
        none.disabled = count === 0;
      };
      updateCount();
      if (!candidates.length) rows.append(el("p", s.noChanges, "hint"));
      for (const candidate of candidates) {
        const kind = reportChangeLabel(viewLabels, candidate.kind, Boolean(candidate.link.reordered));
        const label = [kind, analysis.describe(candidate)].filter(Boolean).join(" · ");
        const { field } = checkbox(label, !excluded.has(candidate.id), (include) => {
          if (include) excluded.delete(candidate.id);
          else excluded.add(candidate.id);
          updateCount();
        });
        rows.append(field);
      }
    }

    updateOptions();
    updatePairs();
  }, undefined);
}
