import type { AppLang, SavedLineageVersion } from "../shared";
import { el, header, modal } from "./extensions";

const labels = {
  en: {
    title: "Import cloud versions", choose: "Choose the saved versions to retrieve.", loading: "Listing versions…",
    none: "No previous versions are available.", status: "Processed", cached: "Local copy only", import: "Retrieve selected versions",
    close: "Close", stop: "Stop", retry: "Retry selected versions", view: "Compare retrieved versions", error: "Could not retrieve",
    "sign-in": "Reopen the cloud document and sign in again.", forbidden: "Access denied",
    missing: "Version no longer available", network: "Connection failed", "too-large": "File is too large",
    invalid: "Could not read the version", conflict: "The version differs from a previous import", cancelled: "Stopped",
  },
  ja: {
    title: "クラウドの過去版を取得", choose: "取得する保存版を選択してください。", loading: "版を一覧にしています…",
    none: "取得できる過去版がありません。", status: "処理済み", cached: "ローカル保管分", import: "選択した版を取得",
    close: "閉じる", stop: "中断", retry: "選択した版を再試行", view: "取得した版を比較", error: "取得できなかった版",
    "sign-in": "クラウド文書を開き直してサインインしてください。", forbidden: "アクセスできません",
    missing: "版が見つかりません", network: "通信に失敗しました", "too-large": "ファイルが上限を超えています",
    invalid: "版を読み取れませんでした", conflict: "以前の取得内容と一致しません", cancelled: "中断しました",
  },
};

export function cloudVersionPickerLabel(lang: AppLang): string { return labels[lang].title; }

export async function pickCloudVersions(lang: AppLang, path: string): Promise<SavedLineageVersion[] | null> {
  const s = labels[lang];
  const describe = (reason: string) => reason in s ? s[reason as keyof typeof s] : s.network;
  let unsubscribe: (() => void) | undefined;
  let active = false;
  try {
    return await modal<SavedLineageVersion[] | null>((dialog, close) => {
      header(dialog, s.title);
      const content = el("section");
      const state = el("p", s.loading, "hint");
      state.setAttribute("role", "status");
      const choices = el("fieldset");
      choices.append(el("legend", s.choose));
      const outcomes = el("ul");
      const actions = el("div", undefined, "lineage-controls");
      const retrieve = el("button", s.import);
      retrieve.disabled = true;
      const stop = el("button", s.stop); stop.disabled = true;
      const view = el("button", s.view); view.disabled = true;
      const cancel = el("button", s.close);
      let choicesReady = false;
      let versions: SavedLineageVersion[] = [];
      const selected = () => [...choices.querySelectorAll<HTMLInputElement>("input:checked")].map(input => input.value);
      const update = () => {
        stop.disabled = !active;
        retrieve.disabled = active || !choicesReady || !selected().length;
        choices.disabled = active;
        view.disabled = active || !versions.some(version => !version.unavailable);
      };
      cancel.onclick = () => close(null);
      view.onclick = () => close(versions);
      stop.onclick = () => void window.app.cancelCloudVersions(path);
      choices.onchange = update;
      retrieve.onclick = () => void (async () => {
        active = true; update();
        outcomes.replaceChildren();
        state.textContent = s.loading;
        try {
          const result = await window.app.importCloudVersions(path, selected());
          versions = result.versions;
          if (result.error) state.textContent = describe(result.error);
          for (const issue of result.failures) outcomes.append(el("li", `${issue.id} · ${describe(issue.reason)}`));
          if (result.failures.length) retrieve.textContent = s.retry;
          if (!result.error && !result.failures.length) state.textContent = `${s.status} ${selected().length}/${selected().length}`;
        } catch {
          state.textContent = s.network;
        } finally {
          active = false; update();
        }
      })();
      actions.append(retrieve, stop, view, cancel);
      content.append(state, choices, outcomes, actions);
      dialog.append(content);
      unsubscribe = window.app.onCloudVersionProgress(event => {
        if (event.path === path) state.textContent = `${s.status} ${event.done}/${event.total}`;
      });
      void window.app.listCloudVersions(path).then(result => {
        state.textContent = result.error ? describe(result.error) : result.items.length ? "" : s.none;
        choicesReady = result.items.length > 0;
        for (const item of result.items) {
          const field = el("label", undefined, "check-row");
          const input = el("input");
          input.type = "checkbox"; input.value = item.id;
          const time = item.modified && Number.isFinite(Date.parse(item.modified)) ? new Date(item.modified).toLocaleString(lang) : "";
          field.append(input, [item.id, time, item.size === undefined ? "" : `${item.size.toLocaleString(lang)} B`, item.cachedOnly ? s.cached : ""].filter(Boolean).join(" · "));
          choices.append(field);
        }
        update();
      }).catch(() => { state.textContent = s.network; });
    }, null);
  } finally {
    unsubscribe?.();
    if (active) await window.app.cancelCloudVersions(path).catch(() => undefined);
  }
}
