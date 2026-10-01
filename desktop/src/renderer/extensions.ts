/**
 * レンダラ側の拡張の扱い: 設定画面の一覧と権限の確認、拡張の分類器をパイプラインに加えること、
 * 解析結果の通知、拡張が求めた UI (ダイアログ・フォーム・パネル) の表示。
 * 拡張から来た文字列は textContent でだけ表示する (HTML として解釈しない)。
 */
import {
  assignLevels,
  Classifier,
  ClassificationResult,
  AnalysisContext,
  FlowResult,
  Highlight,
  PositionedRevisionEvent,
  RuleSet,
  DocxRevisionData,
} from "../../../src/core";
import type {
  AnalysisResult,
  DialogSpec,
  FormSpec,
  FormResult,
  LocalizedText,
  PanelBlock,
  PanelSpec,
  PositionedEvent,
} from "../extension-api";
import type { AppLang, ExtensionListItem, ExtensionUiRequest } from "../shared";
import type { Strings } from "./strings";

export interface ExtUiContext {
  lang(): AppLang;
  strings(): Strings;
  /** パネルが変わったときに呼ぶ (ハイライトのタブを描き直す) */
  onPanelsChanged(): void;
}

let ui: ExtUiContext;
/** "拡張の id\0パネルの id" → パネル */
const panels = new Map<string, { extName: LocalizedText; spec: PanelSpec }>();

export function loc(v: LocalizedText | undefined, lang: AppLang): string {
  if (v === undefined) return "";
  if (typeof v === "string") return v;
  return v[lang] ?? v.en ?? v.ja ?? "";
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, text?: string, cls?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (text !== undefined) e.textContent = text;
  if (cls) e.className = cls;
  return e;
}

// ---------------------------------------------------------------------------
// モーダル (ダイアログ・フォーム・権限の確認)
// ---------------------------------------------------------------------------

let modalQueue: Promise<unknown> = Promise.resolve();

/** モーダルを1つずつ表示する。build は中身を作り、close(value) で閉じる */
function modal<T>(build: (root: HTMLDialogElement, close: (v: T) => void) => void, onCancel: T): Promise<T> {
  const run = () =>
    new Promise<T>((resolve) => {
      const d = document.getElementById("modal") as HTMLDialogElement;
      d.replaceChildren();
      let done = false;
      const close = (v: T) => {
        if (done) return;
        done = true;
        d.close();
        resolve(v);
      };
      d.oncancel = (e) => {
        e.preventDefault();
        close(onCancel);
      };
      build(d, close);
      d.showModal();
    });
  const p = modalQueue.then(run, run);
  modalQueue = p.catch(() => undefined);
  return p;
}

function header(root: HTMLElement, title: string, from?: string): void {
  root.append(el("h3", title));
  if (from) root.append(el("p", from, "from"));
}

function showDialog(extName: LocalizedText, spec: DialogSpec): Promise<{ buttonId: string | null }> {
  const S = ui.strings();
  const lang = ui.lang();
  return modal<{ buttonId: string | null }>((root, close) => {
    header(root, loc(spec.title, lang), S.fromExtension(loc(extName, lang)));
    root.append(el("p", loc(spec.message, lang)));
    const buttons = el("div", undefined, "buttons");
    const specs = spec.buttons?.length ? spec.buttons : [{ id: "ok", label: S.ok, primary: true }];
    for (const b of specs) {
      const btn = el("button", loc(b.label, lang), b.primary ? "primary" : undefined);
      btn.onclick = () => close({ buttonId: b.id });
      buttons.append(btn);
    }
    root.append(buttons);
  }, { buttonId: null });
}

function openForm(extName: LocalizedText, spec: FormSpec): Promise<FormResult | null> {
  const S = ui.strings();
  const lang = ui.lang();
  return modal<FormResult | null>((root, close) => {
    header(root, loc(spec.title, lang), S.fromExtension(loc(extName, lang)));
    if (spec.description) root.append(el("p", loc(spec.description, lang)));
    const inputs = new Map<string, HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>();
    for (const f of spec.fields ?? []) {
      const wrap = el("label", undefined, f.type === "checkbox" ? "field check" : "field");
      const label = el("span", loc(f.label, lang));
      let input: HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement;
      if (f.type === "textarea") {
        input = el("textarea");
        input.value = f.default ?? "";
        if (f.maxLength) input.maxLength = f.maxLength;
      } else if (f.type === "select") {
        input = el("select");
        for (const o of f.options ?? []) {
          const opt = el("option", loc(o.label, lang));
          opt.value = o.value;
          input.append(opt);
        }
        if (f.default !== undefined) input.value = f.default;
      } else {
        input = el("input");
        input.type = f.type === "number" ? "number" : f.type === "checkbox" ? "checkbox" : "text";
        if (f.type === "checkbox") input.checked = !!f.default;
        else if (f.default !== undefined) input.value = String(f.default);
        if (f.type === "number") {
          if (f.min !== undefined) input.min = String(f.min);
          if (f.max !== undefined) input.max = String(f.max);
        }
        if (f.type === "text" && f.maxLength) input.maxLength = f.maxLength;
      }
      inputs.set(f.id, input);
      if (f.type === "checkbox") wrap.append(input, label);
      else wrap.append(label, input);
      root.append(wrap);
    }
    const error = el("p", "", "error");
    error.hidden = true;
    const buttons = el("div", undefined, "buttons");
    const cancel = el("button", S.cancel);
    cancel.onclick = () => close(null);
    const submit = el("button", loc(spec.submitLabel, lang) || S.ok, "primary");
    submit.onclick = () => {
      const values: FormResult = {};
      for (const f of spec.fields ?? []) {
        const input = inputs.get(f.id)!;
        if (f.type === "checkbox") values[f.id] = (input as HTMLInputElement).checked;
        else if (f.type === "number") {
          const n = parseFloat(input.value);
          if (input.value !== "" && Number.isFinite(n)) values[f.id] = n;
        } else values[f.id] = input.value;
        if ("required" in f && f.required && (values[f.id] === undefined || values[f.id] === "")) {
          error.textContent = S.formRequired(loc(f.label, lang));
          error.hidden = false;
          return;
        }
      }
      close(values);
    };
    buttons.append(cancel, submit);
    root.append(error, buttons);
  }, null);
}

/** 拡張を有効にする前に、求める権限を示して確認する */
export function confirmPermissions(item: ExtensionListItem): Promise<boolean> {
  const S = ui.strings();
  const lang = ui.lang();
  return modal<boolean>((root, close) => {
    header(root, S.extConfirmTitle(loc(item.manifest?.name ?? item.id, lang)));
    root.append(el("p", S.extConfirmBody));
    root.append(permissionList(item));
    const buttons = el("div", undefined, "buttons");
    const cancel = el("button", S.cancel);
    cancel.onclick = () => close(false);
    const ok = el("button", S.extConfirm, "primary");
    ok.onclick = () => close(true);
    buttons.append(cancel, ok);
    root.append(buttons);
  }, false);
}

export function permissionList(item: ExtensionListItem): HTMLUListElement {
  const S = ui.strings();
  const p = item.manifest?.permissions ?? {};
  const ul = el("ul");
  const items: string[] = [];
  if (p.ui?.length) items.push(S.extPermUi(p.ui.join(", ")));
  if (p.storage) items.push(S.extPermStorage);
  if (p.network?.length) items.push(S.extPermNetwork(p.network.join(", ")));
  if (p.documentText) items.push(S.extPermText);
  for (const t of items.length ? items : [S.extPermNone]) ul.append(el("li", t));
  return ul;
}

// ---------------------------------------------------------------------------
// パネル
// ---------------------------------------------------------------------------

export function clearPanels(): void {
  panels.clear();
}

/** パネル・ポップアップのブロックを要素にする (文字列は textContent でだけ表示する) */
export function renderBlocks(blocks: PanelBlock[]): HTMLElement[] {
  const lang = ui.lang();
  const out: HTMLElement[] = [];
  for (const b of blocks) {
    if (b.type === "heading") out.push(el("h4", loc(b.text, lang)));
    else if (b.type === "text") out.push(el("p", loc(b.text, lang)));
    else if (b.type === "list") {
      const ul = el("ul");
      for (const i of b.items ?? []) ul.append(el("li", loc(i, lang)));
      out.push(ul);
    } else if (b.type === "keyValue") {
      const t = el("table");
      for (const r of b.rows ?? []) {
        const tr = el("tr");
        tr.append(el("th", loc(r.key, lang)), el("td", String(r.value)));
        t.append(tr);
      }
      out.push(t);
    } else if (b.type === "table") {
      const t = el("table");
      const head = el("tr");
      for (const c of b.columns ?? []) head.append(el("th", loc(c, lang)));
      t.append(head);
      for (const r of b.rows ?? []) {
        const tr = el("tr");
        for (const c of r) tr.append(el("td", String(c)));
        t.append(tr);
      }
      out.push(t);
    }
  }
  return out;
}

/** 画面の表示言語と文言 (figures.ts から使う) */
export function uiContext(): ExtUiContext {
  return ui;
}

/** ハイライトのタブの下に、拡張のパネルを描く */
export function renderPanels(container: HTMLElement): void {
  container.querySelectorAll(".ext-panel, .ext-panels-title").forEach((n) => n.remove());
  if (panels.size === 0) return;
  const S = ui.strings();
  const lang = ui.lang();
  container.append(el("h2", S.extPanels, "ext-panels-title"));
  for (const { extName, spec } of panels.values()) {
    const box = el("section", undefined, "ext-panel");
    box.append(el("h3", loc(spec.title, lang)), el("p", S.fromExtension(loc(extName, lang)), "hint"));
    box.append(...renderBlocks(spec.blocks ?? []));
    container.append(box);
  }
}

function prettyJson(json: string): string {
  try {
    return JSON.stringify(JSON.parse(json), null, 2);
  } catch {
    return json;
  }
}

function originOf(url: string): string {
  try {
    return new URL(url).origin;
  } catch {
    return url;
  }
}

/** 拡張の送信の前に、宛先と本文を見せて確認する */
function confirmSend(extName: LocalizedText, spec: { url: string; body: string }): Promise<boolean> {
  const S = ui.strings();
  const lang = ui.lang();
  return modal<boolean>((root, close) => {
    root.classList.add("wide");
    header(root, S.sendConfirmTitle, S.fromExtension(loc(extName, lang)));
    root.append(el("p", S.sendConfirmBody(originOf(spec.url))));
    root.append(el("strong", S.sendTo), el("p", spec.url, "mono"));
    root.append(el("strong", S.sendData), el("pre", prettyJson(spec.body)));
    const buttons = el("div", undefined, "buttons");
    const no = el("button", S.sendDecline);
    no.onclick = () => close(false);
    const yes = el("button", S.sendAllow, "primary");
    yes.onclick = () => close(true);
    buttons.append(no, yes);
    root.append(buttons);
    // 既定のボタンは「送信しない」(Enter で誤って送らない)
    setTimeout(() => no.focus(), 0);
  }, false).finally(() => (document.getElementById("modal") as HTMLDialogElement).classList.remove("wide"));
}

/** 図の注釈のリンクを開く前に、URL 全体を見せて確認する (URL にデータを含めて外へ出せるため) */
function confirmOpen(extName: LocalizedText, url: string): Promise<boolean> {
  const S = ui.strings();
  const lang = ui.lang();
  return modal<boolean>((root, close) => {
    header(root, S.openConfirmTitle, S.fromExtension(loc(extName, lang)));
    root.append(el("p", S.openConfirmBody(originOf(url))));
    root.append(el("pre", url));
    const buttons = el("div", undefined, "buttons");
    const no = el("button", S.openDecline);
    no.onclick = () => close(false);
    const yes = el("button", S.openAllow, "primary");
    yes.onclick = () => close(true);
    buttons.append(no, yes);
    root.append(buttons);
    setTimeout(() => no.focus(), 0);
  }, false);
}

/** 拡張の送信履歴を表示する */
export async function showSendLog(item: ExtensionListItem): Promise<void> {
  const S = ui.strings();
  const lang = ui.lang();
  const entries = (await window.app.extensions.sendLog(item.id)).slice().reverse();
  await modal<void>((root, close) => {
    root.classList.add("wide");
    header(root, S.sendLogTitle(loc(item.manifest?.name ?? item.id, lang)));
    root.append(el("p", S.sendLogNote, "hint"));
    const list = el("div", undefined, "sendlog");
    if (entries.length === 0) list.append(el("p", S.sendLogEmpty, "hint"));
    for (const e of entries) {
      const d = el("details");
      const when = new Date(e.time).toLocaleString(lang === "ja" ? "ja-JP" : "en-US");
      const status = `${S.sendStatus[e.status] ?? e.status}${e.httpStatus ? ` (HTTP ${e.httpStatus})` : ""}`;
      d.append(el("summary", `${when} — ${status} — ${e.url}`));
      if (e.error) d.append(el("p", e.error, "hint"));
      d.append(el("pre", prettyJson(e.body) + (e.truncated ? "\n…" : "")));
      list.append(d);
    }
    root.append(list);
    const buttons = el("div", undefined, "buttons");
    const ok = el("button", S.close, "primary");
    ok.onclick = () => close();
    buttons.append(ok);
    root.append(buttons);
  }, undefined);
  (document.getElementById("modal") as HTMLDialogElement).classList.remove("wide");
}

export function initExtensionUi(ctx: ExtUiContext): void {
  ui = ctx;
  window.app.extensions.onUiRequest(async (req: ExtensionUiRequest) => {
    if (req.kind === "confirmSend") return confirmSend(req.ext.name, req.spec as { url: string; body: string });
    if (req.kind === "confirmOpen") return confirmOpen(req.ext.name, (req.spec as { url: string }).url);
    if (req.kind === "dialog") return showDialog(req.ext.name, req.spec as DialogSpec);
    if (req.kind === "form") return openForm(req.ext.name, req.spec as FormSpec);
    const spec = req.spec as PanelSpec;
    panels.set(`${req.ext.id}\0${spec.id}`, { extName: req.ext.name, spec });
    ui.onPanelsChanged();
    return null;
  });
}

// ---------------------------------------------------------------------------
// 分類器と解析結果
// ---------------------------------------------------------------------------

export function toJsonEvent(e: PositionedRevisionEvent): PositionedEvent {
  return {
    type: e.type,
    id: e.id,
    date: e.date ? e.date.toISOString() : null,
    author: e.author ?? null,
    move: e.move,
    moveName: e.moveName ?? null,
    chars: e.chars,
    text: e.text,
    includesParaMark: e.includesParaMark,
    paraModelIndices: e.paraModelIndices,
    final: e.final,
    finalChars: e.finalChars,
    atEdit: e.atEdit ?? null,
  };
}

/** 有効な拡張の分類器を、アプリの分類器のパイプラインに加えられる形にする */
export async function extensionClassifiers(onError: (msg: string) => void): Promise<Classifier[]> {
  const regs = await window.app.extensions.registrations();
  const out: Classifier[] = [];
  for (const r of regs) {
    const categories = r.categories.map((c) => ({
      id: c.id,
      role: "highlight" as const,
      color: c.color,
      label: c.label,
      priority: c.priority,
      pattern: c.pattern,
    }));
    const known = new Set(regs.flatMap((x) => x.categories.map((c) => c.id)));
    for (const c of r.classifiers) {
      out.push({
        id: c.id,
        version: c.version,
        categories,
        async classify(ctx: AnalysisContext): Promise<Highlight[]> {
          const res = await window.app.extensions.classify(r.extId, c.id, {
            positioned: ctx.positioned.map(toJsonEvent),
            sessions: ctx.sessions.map((s) => ({ start: s.start.toISOString(), end: s.end.toISOString() })),
          });
          if (res.error) onError(ui.strings().extClassifierFailed(c.id, res.error));
          const byId = new Map(ctx.positioned.filter((e) => e.type === "ins").map((e) => [e.id, e]));
          return res.highlights
            .filter((h) => known.has(h.categoryId) && Array.isArray(h.eventIds))
            .map((h): Highlight => {
              const evs = h.eventIds.map((id) => byId.get(String(id))).filter((e): e is PositionedRevisionEvent => !!e);
              const start = evs.reduce((a, e) => (!a || e.final.start.docOffset < a.docOffset ? e.final.start : a), evs[0]?.final.start);
              const end = evs.reduce((a, e) => (!a || e.final.end.docOffset > a.docOffset ? e.final.end : a), evs[0]?.final.end);
              return {
                classifierId: c.id,
                categoryId: h.categoryId,
                eventIds: h.eventIds.map(String),
                timeRange: { start: new Date(h.start), end: new Date(h.end) },
                docRange: start && end ? { start, end } : undefined,
                features: h.features,
                reason: h.reason,
              };
            });
        },
      });
    }
  }
  return out;
}

/** 拡張に渡す解析結果を作る (本文は、権限の無い拡張にはメインプロセスが渡さない) */
export function buildAnalysisResult(args: {
  file: { name: string; mtime: string };
  data: DocxRevisionData;
  ctx: AnalysisContext;
  rules: RuleSet;
  classification: ClassificationResult;
  flow?: FlowResult;
  finalText: string;
  integrity: AnalysisResult["integrity"];
}): AnalysisResult {
  const { data, ctx, rules, classification, flow } = args;
  const windows = ctx.windowsFor(rules.window);
  const levels = assignLevels(rules, windows);
  const { events: _events, ...summary } = data;
  return {
    file: args.file,
    revisions: { ...summary, eventCount: data.events.length },
    positioned: ctx.positioned.map(toJsonEvent),
    sessions: (flow?.sessions ?? []).map((s) => ({
      start: s.start.toISOString(),
      end: s.end.toISOString(),
      insChars: s.insChars,
      delChars: s.delChars,
      bulkInsChars: s.bulkInsChars,
      movedChars: s.movedChars,
    })),
    ruleSet: rules.ruleSet,
    windows: windows.windows.map((w, i) => ({
      index: w.index,
      start: w.start.toISOString(),
      end: w.end.toISOString(),
      authors: w.authors,
      eventIds: w.eventIds,
      range: w.range,
      features: { ...w.features },
      level: levels[i]?.id ?? null,
    })),
    highlights: classification.highlights.map((h) => ({
      classifierId: h.classifierId,
      categoryId: h.categoryId,
      eventIds: h.eventIds,
      start: h.timeRange.start.toISOString(),
      end: h.timeRange.end.toISOString(),
      features: h.features,
      reason: h.reason,
    })),
    finalText: args.finalText,
    integrity: args.integrity,
  };
}
