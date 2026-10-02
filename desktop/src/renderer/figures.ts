/**
 * 図 (chart / flow の SVG) の部分への、拡張の注釈。
 *  - 有効な拡張の注釈器 (registerFigureAnnotator) に図の部分の一覧を渡し、注釈を集める
 *  - アプリの中: 注釈のある部分にマウスを重ねるとポップアップ (説明・ブロック・リンク) を出し、
 *    クリックでリンクを開く (確認と送信履歴への記録はメインプロセスが行う)
 *  - 書き出す SVG: 説明を <title>、リンクを <a href> として埋め込む (ブラウザで開いても動く)
 */
import { FigureTarget, isSafeHref, ResolvedAnnotation, resolveAnnotations } from "../../../src/core";
import type { AnalysisResult, FigureAnnotationSpec, LocalizedText } from "../extension-api";
import type { AppLang } from "../shared";
import { loc, renderBlocks, uiContext } from "./extensions";

export interface AnnotationEntry {
  extId: string;
  extName: LocalizedText;
  spec: FigureAnnotationSpec;
}

/** 図 → 部分のキー → 注釈 */
export type FigureAnnotations = Record<"chart" | "flow", Map<string, AnnotationEntry[]>>;

/** 有効な拡張の注釈器から注釈を集める (失敗した注釈器はエラーとして知らせ、ほかは続ける) */
export async function collectFigureAnnotations(
  targets: Record<"chart" | "flow", FigureTarget[]>,
  analysis: AnalysisResult,
  onError: (msg: string) => void
): Promise<FigureAnnotations> {
  const out: FigureAnnotations = { chart: new Map(), flow: new Map() };
  const regs = await window.app.extensions.registrations();
  for (const r of regs) {
    for (const an of r.annotators) {
      for (const figure of ["chart", "flow"] as const) {
        if (targets[figure].length === 0) continue;
        const res = await window.app.extensions.annotateFigure(r.extId, an.id, {
          figure,
          targets: targets[figure],
          analysis,
        });
        if (res.error) onError(uiContext().strings().figureAnnotatorFailed(an.id, res.error));
        for (const spec of res.annotations) {
          const list = out[figure].get(spec.target) ?? [];
          list.push({ extId: r.extId, extName: r.name, spec });
          out[figure].set(spec.target, list);
        }
      }
    }
  }
  return out;
}

/** 書き出す SVG に埋め込む注釈 (説明は tooltip、無ければポップアップの見出し) */
export function annotationsForExport(map: Map<string, AnnotationEntry[]>, lang: AppLang): Map<string, ResolvedAnnotation> {
  const list = [...map.values()].flat().map((e) => ({
    target: e.spec.target,
    tooltip: e.spec.tooltip ?? e.spec.popup?.title,
    href: e.spec.href,
  }));
  return resolveAnnotations(list, lang);
}

let popup: HTMLDivElement | undefined;
let pinned = false;

function popupEl(): HTMLDivElement {
  if (!popup) {
    popup = document.createElement("div");
    popup.className = "figure-popup";
    popup.hidden = true;
    popup.setAttribute("role", "tooltip");
    document.body.append(popup);
    document.addEventListener("click", (e) => {
      if (pinned && popup && !popup.contains(e.target as Node)) hidePopup();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") hidePopup();
    });
  }
  return popup;
}

function hidePopup(): void {
  pinned = false;
  if (popup) popup.hidden = true;
  for (const el of document.querySelectorAll(".dra-hover")) el.classList.remove("dra-hover");
}

function originOf(url: string): string {
  try {
    return new URL(url).origin;
  } catch {
    return url;
  }
}

function openLink(e: AnnotationEntry): void {
  if (e.spec.href) void window.app.extensions.openLink(e.extId, e.spec.href);
}

function fillPopup(entries: AnnotationEntry[], target?: FigureTarget): void {
  const { strings, lang } = uiContext();
  const S = strings();
  const p = popupEl();
  p.replaceChildren();
  // 段落・帯なら、その段落の冒頭 (約 40 文字) とセクションを先頭に示す (本文はアプリの中だけで表示する)
  if (target?.excerpt) {
    const head = document.createElement("section");
    head.className = "excerpt";
    const q = document.createElement("p");
    q.textContent = `「${target.excerpt}」`;
    head.append(q);
    if (target.section) {
      const sec = document.createElement("p");
      sec.className = "from";
      sec.textContent = S.figureSection(target.section);
      head.append(sec);
    }
    p.append(head);
  }
  for (const e of entries) {
    const sec = document.createElement("section");
    const from = document.createElement("p");
    from.className = "from";
    from.textContent = S.fromExtension(loc(e.extName, lang()));
    sec.append(from);
    if (e.spec.popup?.title) {
      const h = document.createElement("h4");
      h.textContent = loc(e.spec.popup.title, lang());
      sec.append(h);
    }
    if (e.spec.tooltip) {
      const t = document.createElement("p");
      t.textContent = loc(e.spec.tooltip, lang());
      sec.append(t);
    }
    if (e.spec.popup?.blocks) sec.append(...renderBlocks(e.spec.popup.blocks));
    if (e.spec.href && isSafeHref(e.spec.href)) {
      const b = document.createElement("button");
      b.className = "link";
      // 元の文書 (OneDrive / SharePoint) の該当箇所へのリンクなら、そう分かる名前にする
      b.textContent =
        target?.docLink && e.spec.href === target.docLink
          ? S.figureOpenSource(target.section)
          : S.figureOpenLink(originOf(e.spec.href));
      b.onclick = (ev) => {
        ev.stopPropagation();
        openLink(e);
      };
      sec.append(b);
    }
    p.append(sec);
  }
}

function place(x: number, y: number): void {
  const p = popupEl();
  p.hidden = false;
  const pad = 14;
  const w = p.offsetWidth;
  const h = p.offsetHeight;
  p.style.left = `${Math.min(x + pad, window.innerWidth - w - 8)}px`;
  p.style.top = `${y + pad + h > window.innerHeight ? Math.max(8, y - h - pad) : y + pad}px`;
}

/** 図の要素に、注釈のポップアップとリンクを結びつける */
export function bindFigureAnnotations(pane: HTMLElement, map: Map<string, AnnotationEntry[]>, targets: FigureTarget[] = []): void {
  if (map.size === 0) return;
  const targetOf = new Map(targets.map((t) => [t.key, t]));
  const S = uiContext().strings();
  for (const el of pane.querySelectorAll<SVGElement>("[data-target]")) {
    const entries = map.get(el.dataset.target!);
    if (!entries) continue;
    el.classList.add("dra-annotated");
    const links = entries.filter((e) => e.spec.href);
    el.setAttribute("aria-label", entries.map((e) => loc(e.spec.tooltip ?? e.spec.popup?.title, uiContext().lang())).join(" / ") || S.figureAnnotated);
    el.addEventListener("mouseenter", (ev) => {
      if (pinned) return;
      el.classList.add("dra-hover");
      fillPopup(entries, targetOf.get(el.dataset.target!));
      place(ev.clientX, ev.clientY);
    });
    el.addEventListener("mousemove", (ev) => {
      if (!pinned) place(ev.clientX, ev.clientY);
    });
    el.addEventListener("mouseleave", () => {
      if (!pinned) hidePopup();
    });
    el.addEventListener("click", (ev) => {
      ev.stopPropagation();
      if (links.length === 1) {
        hidePopup();
        openLink(links[0]);
      } else {
        // リンクが無い・複数あるときは、ポップアップを固定してボタンから選べるようにする
        hidePopup();
        fillPopup(entries, targetOf.get(el.dataset.target!));
        place(ev.clientX, ev.clientY);
        pinned = true;
      }
    });
  }
}
