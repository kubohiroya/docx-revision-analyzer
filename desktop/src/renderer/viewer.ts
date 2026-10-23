/**
 * 図 (chart / flow の SVG) のプレビュー。
 *  - ドラッグで移動、Ctrl / ⌘ + ホイール (トラックパッドのピンチ) でカーソルの位置を中心に拡大縮小。
 *    倍率の表示と「全体」「100%」のボタン。倍率は図ごとに保ち、同じ文書を描き直しても維持する (#54)
 *  - ミニマップ: 図全体の縮小表示と、いまの表示範囲。クリック・ドラッグで移動、ダブルクリックで全体表示、
 *    枠のドラッグで大きさの変更、マウスを重ねてカーソルキーで四隅へ移動 (#55)
 *  - flow: ページをクリックして強調し、右の画面に大きく表示する (スプリッターで幅を変えられる)。
 *    カーソルキーで強調するページを動かし、ページ以外のクリックか Esc で解除する (#56)
 *  - flow の右の画面は「模式図 / 保存時のページ」を切り替えられる。保存時のページは、その列の時点で保存した版の PDF を
 *    メインプロセスが画像にしたもの (#59)。↑ ↓ は保存時のページの中で動く
 *
 * 拡大縮小は SVG の表示の大きさ (width / height) を変えて行う (ベクターのまま描き直すのでぼやけない)。表示枠は
 * ふつうのスクロールの領域なので、scrollIntoView や注釈のポップアップの位置はそのまま働く。
 * 強調の枠やミニマップは画面の中だけのもので、書き出す SVG / PNG (別に描き直す) には含まれない。
 */
import type { Strings } from "./strings";

export type FigureKind = "chart" | "flow";
export type Corner = "tl" | "tr" | "bl" | "br";

export interface MinimapPrefs {
  /** ミニマップの幅 (px) */
  width: number;
  corner: Corner;
}

export interface ViewerPrefs {
  minimap?: MinimapPrefs;
  /** flow の分割画面の左の幅 (0〜1) */
  split?: number;
}

/** 列の時点で保存した版の、保存時のページの状態 */
export interface SavedPageInfo {
  /** 版の説明 (保存した日時) */
  label: string;
  /** off: 設定で作らないことにしている */
  state: "none" | "pending" | "done" | "failed" | "off";
  pageCount?: number;
  error?: string;
}

export interface ViewerHooks {
  strings: () => Strings;
  prefs: () => ViewerPrefs;
  savePrefs: (p: ViewerPrefs) => void;
  /** flow: 列 (どの時点の文書か) の見出し */
  columnTitle?: (column: number) => string;
  /** flow: 右の画面に表示したページ (図の複製) に、注釈などを結びつける */
  onDetail?: (body: HTMLElement) => void;
  /** flow: 保存時のページ (無ければ「模式図」だけ) */
  saved?: {
    /** 列の時点で保存した版 (無ければ undefined) */
    info(column: number): SavedPageInfo | undefined;
    /** 保存時のページを作る */
    request(column: number): void;
    /** 保存時のページ (0 から) を maxWidth × maxHeight (画素) に収まる画像にする */
    render(column: number, page: number, maxWidth: number, maxHeight: number): Promise<{ dataUrl: string; pageCount: number }>;
  };
}

const MIN_ZOOM = 0.05;
const MAX_ZOOM = 8;
const DRAG_THRESHOLD = 4;
const MINIMAP_MIN = 100;
const MINIMAP_MAX = 480;
const MINIMAP_MARGIN = 12;
const EDGE = 7;

function clamp(v: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, v));
}

function div(className: string): HTMLDivElement {
  const el = document.createElement("div");
  el.className = className;
  return el;
}

function button(text: string, title: string, onClick: () => void): HTMLButtonElement {
  const b = document.createElement("button");
  b.textContent = text;
  b.title = title;
  b.onclick = (e) => {
    e.stopPropagation();
    onClick();
  };
  return b;
}

/** ポインタを要素に結びつける (ドラッグが要素の外へ出ても続けるため。結びつけられなくても操作は続ける) */
function capture(el: Element, pointerId: number): void {
  try {
    el.setPointerCapture(pointerId);
  } catch {
    // 合成したイベントなど
  }
}

/** 入力欄にフォーカスがあるときは、カーソルキーを図の操作に使わない */
function typing(e: KeyboardEvent): boolean {
  const t = e.target as HTMLElement | null;
  return !!t && (t.isContentEditable || ["INPUT", "SELECT", "TEXTAREA"].includes(t.tagName));
}

export class FigureViewer {
  /** 図の表示枠 (スクロールの領域)。注釈の結びつけや位置の検索はこの中で行う */
  readonly viewport: HTMLDivElement;
  private readonly left: HTMLDivElement;
  private readonly zoomLabel: HTMLSpanElement;
  private readonly bar: HTMLDivElement;
  private readonly buttons: Record<"out" | "in" | "fit" | "actual", HTMLButtonElement>;
  private readonly minimap: Minimap;
  private readonly splitter?: HTMLDivElement;
  private readonly detail?: HTMLDivElement;
  private readonly detailHead?: HTMLDivElement;
  private readonly detailTitle?: HTMLSpanElement;
  private readonly modeButtons?: { schematic: HTMLButtonElement; saved: HTMLButtonElement };
  private readonly detailBody?: HTMLDivElement;
  /** 右の画面: 模式図 / 保存時のページ */
  private detailMode: "schematic" | "saved" = "schematic";
  private savedPage = 0;
  private savedPageCount?: number;
  private savedRequest = 0;
  private savedResize?: number;
  private svg?: SVGSVGElement;
  private natural = { w: 0, h: 0 };
  private zoomValue = 1;
  /** 表示枠の大きさが決まってから全体表示にする (タブが隠れている間は大きさが 0) */
  private pendingFit = false;
  private suppressClick = false;
  private selected?: { column: number; page: number };
  /** 右の画面のページの縦横比 (幅 / 高さ) */
  private detailAspect = 1;

  constructor(
    readonly host: HTMLElement,
    readonly kind: FigureKind,
    private readonly hooks: ViewerHooks
  ) {
    const S = hooks.strings();
    host.classList.add("figure-host");
    host.replaceChildren();
    this.left = div("figure-left");
    this.viewport = div("figure viewport");
    this.viewport.title = "";
    this.bar = div("zoombar");
    this.zoomLabel = document.createElement("span");
    this.zoomLabel.className = "zoom-label";
    this.buttons = {
      out: button("−", S.zoomOut, () => this.setZoom(this.zoomValue / 1.25)),
      in: button("+", S.zoomIn, () => this.setZoom(this.zoomValue * 1.25)),
      fit: button(S.zoomFit, S.zoomFit, () => this.fit()),
      actual: button("100%", S.zoom100, () => this.setZoom(1)),
    };
    this.bar.append(this.buttons.out, this.zoomLabel, this.buttons.in, this.buttons.fit, this.buttons.actual);
    this.minimap = new Minimap(this, hooks);
    this.left.append(this.viewport, this.bar, this.minimap.el);
    host.append(this.left);
    if (kind === "flow") {
      this.splitter = div("splitter");
      this.splitter.title = S.splitterTitle;
      this.detail = div("page-detail");
      this.detailHead = div("page-detail-head");
      this.detailTitle = document.createElement("span");
      this.detailTitle.className = "page-detail-title";
      const modes = div("page-detail-modes");
      this.modeButtons = {
        schematic: button(S.detailSchematic, S.detailSchematic, () => this.setDetailMode("schematic")),
        saved: button(S.detailSaved, S.detailSaved, () => this.setDetailMode("saved")),
      };
      modes.append(this.modeButtons.schematic, this.modeButtons.saved);
      modes.hidden = !hooks.saved;
      this.detailHead.append(this.detailTitle, modes);
      this.detailBody = div("page-detail-body");
      this.detail.append(this.detailHead, this.detailBody);
      host.append(this.splitter, this.detail);
      this.initSplitter();
      new ResizeObserver(() => {
        if (this.detailMode === "schematic") this.fitDetail();
        else {
          window.clearTimeout(this.savedResize);
          this.savedResize = window.setTimeout(() => this.renderSaved(), 150);
        }
      }).observe(this.detailBody);
    }
    this.initPanZoom();
    this.applyStrings();
    new ResizeObserver(() => this.onResize()).observe(this.viewport);
    document.addEventListener("keydown", (e) => this.onKey(e));
  }

  /** 表示言語が変わったとき */
  applyStrings(): void {
    const S = this.hooks.strings();
    this.bar.title = S.zoomHint;
    this.buttons.out.title = S.zoomOut;
    this.buttons.in.title = S.zoomIn;
    this.buttons.fit.textContent = S.zoomFit;
    this.buttons.fit.title = S.zoomFit;
    this.buttons.actual.title = S.zoom100;
    this.minimap.el.title = S.minimapHint;
    if (this.splitter) this.splitter.title = S.splitterTitle;
    if (this.detailHead) this.detailHead.title = S.pageDetailHint;
    if (this.modeButtons) {
      this.modeButtons.schematic.textContent = this.modeButtons.schematic.title = S.detailSchematic;
      this.modeButtons.saved.textContent = S.detailSaved;
    }
    this.refreshDetail();
  }

  get zoom(): number {
    return this.zoomValue;
  }

  get naturalSize(): { w: number; h: number } {
    return this.natural;
  }

  /** 表示枠が見えていて、大きさが決まっているか */
  private get visible(): boolean {
    return this.viewport.clientWidth > 0 && this.viewport.clientHeight > 0;
  }

  /**
   * 図を表示する。sameDocument なら倍率・表示位置・ページの強調を保つ (設定の変更などでの描き直し)。
   * svg が無ければ html (案内) を表示する
   */
  setContent(svg: string | undefined, html = "", sameDocument = false): void {
    const keep = sameDocument && !!this.svg;
    const scroll = { left: this.viewport.scrollLeft, top: this.viewport.scrollTop };
    const selected = keep ? this.selected : undefined;
    this.clearSelection();
    this.viewport.innerHTML = svg ?? html;
    this.svg = this.viewport.querySelector("svg") ?? undefined;
    if (this.svg) {
      this.natural = {
        w: parseFloat(this.svg.getAttribute("width") ?? "") || this.svg.viewBox.baseVal.width || 1,
        h: parseFloat(this.svg.getAttribute("height") ?? "") || this.svg.viewBox.baseVal.height || 1,
      };
      if (keep) {
        this.applySize();
        this.viewport.scrollLeft = scroll.left;
        this.viewport.scrollTop = scroll.top;
      } else {
        this.zoomValue = 1;
        this.applySize();
        this.viewport.scrollLeft = 0;
        this.viewport.scrollTop = 0;
        // 既定は全体表示 (ただし 100% より大きくはしない)
        if (this.visible) this.fit(true);
        else this.pendingFit = true;
      }
      if (selected) this.selectPage(selected.column, selected.page, false, true);
    }
    this.minimap.setImage(svg);
    this.updateLabel();
  }

  /** 図を表示枠に収まる倍率にする。noUpscale なら 100% より大きくしない */
  fit(noUpscale = false): void {
    if (!this.svg) return;
    if (!this.visible) {
      this.pendingFit = true;
      return;
    }
    this.pendingFit = false;
    const pad = 8;
    const z = Math.min((this.viewport.clientWidth - pad) / this.natural.w, (this.viewport.clientHeight - pad) / this.natural.h);
    this.setZoom(noUpscale ? Math.min(1, z) : z);
    this.viewport.scrollLeft = 0;
    this.viewport.scrollTop = 0;
  }

  /** 倍率を変える。anchor (表示枠の中の位置) の下の図の点が動かないようにする。既定は表示枠の中心 */
  setZoom(z: number, anchor?: { x: number; y: number }): void {
    if (!this.svg) return;
    const next = clamp(z, MIN_ZOOM, MAX_ZOOM);
    const a = anchor ?? { x: this.viewport.clientWidth / 2, y: this.viewport.clientHeight / 2 };
    const cx = (this.viewport.scrollLeft + a.x) / this.zoomValue;
    const cy = (this.viewport.scrollTop + a.y) / this.zoomValue;
    this.zoomValue = next;
    this.applySize();
    this.viewport.scrollLeft = cx * next - a.x;
    this.viewport.scrollTop = cy * next - a.y;
    this.updateLabel();
    this.minimap.update();
  }

  /** 表示範囲の中心を、図の点 (倍率 1 の座標) に合わせる */
  centerOn(x: number, y: number): void {
    this.viewport.scrollLeft = x * this.zoomValue - this.viewport.clientWidth / 2;
    this.viewport.scrollTop = y * this.zoomValue - this.viewport.clientHeight / 2;
  }

  /** いまの表示範囲 (倍率 1 の座標) */
  visibleRect(): { x: number; y: number; w: number; h: number } {
    const z = this.zoomValue;
    return {
      x: this.viewport.scrollLeft / z,
      y: this.viewport.scrollTop / z,
      w: this.viewport.clientWidth / z,
      h: this.viewport.clientHeight / z,
    };
  }

  /** 図全体が表示枠に収まっているか */
  fitsEntirely(): boolean {
    return this.natural.w * this.zoomValue <= this.viewport.clientWidth + 1 && this.natural.h * this.zoomValue <= this.viewport.clientHeight + 1;
  }

  /** タブを表示したとき */
  shown(): void {
    this.onResize();
  }

  private applySize(): void {
    if (!this.svg) return;
    this.svg.style.width = `${this.natural.w * this.zoomValue}px`;
    this.svg.style.height = `${this.natural.h * this.zoomValue}px`;
  }

  private updateLabel(): void {
    this.zoomLabel.textContent = this.svg ? `${Math.round(this.zoomValue * 100)}%` : "";
    this.left.classList.toggle("empty", !this.svg);
  }

  private onResize(): void {
    if (this.pendingFit && this.visible) this.fit(true);
    this.minimap.layout();
    this.minimap.update();
  }

  // ---- 移動と拡大縮小 ---------------------------------------------------------

  private initPanZoom(): void {
    const vp = this.viewport;
    vp.addEventListener(
      "wheel",
      (e) => {
        if (!(e.ctrlKey || e.metaKey) || !this.svg) return;
        e.preventDefault();
        const r = vp.getBoundingClientRect();
        const delta = e.deltaMode === 1 ? e.deltaY * 33 : e.deltaY;
        this.setZoom(this.zoomValue * Math.exp(-delta * 0.0015), { x: e.clientX - r.left, y: e.clientY - r.top });
      },
      { passive: false }
    );
    let drag: { x: number; y: number; left: number; top: number; id: number; moved: boolean } | undefined;
    vp.addEventListener("pointerdown", (e) => {
      if (e.button !== 0 || !this.svg) return;
      // スクロールバーの上は除く
      if (e.target === vp && (e.offsetX >= vp.clientWidth || e.offsetY >= vp.clientHeight)) return;
      drag = { x: e.clientX, y: e.clientY, left: vp.scrollLeft, top: vp.scrollTop, id: e.pointerId, moved: false };
    });
    vp.addEventListener("pointermove", (e) => {
      if (!drag || e.pointerId !== drag.id) return;
      const dx = e.clientX - drag.x;
      const dy = e.clientY - drag.y;
      if (!drag.moved && Math.hypot(dx, dy) < DRAG_THRESHOLD) return;
      if (!drag.moved) {
        drag.moved = true;
        capture(vp, e.pointerId);
        vp.classList.add("panning");
      }
      vp.scrollLeft = drag.left - dx;
      vp.scrollTop = drag.top - dy;
    });
    const end = (e: PointerEvent) => {
      if (!drag || e.pointerId !== drag.id) return;
      if (drag.moved) {
        // ドラッグの後のクリックは、ページの指定や注釈のクリックとして扱わない
        this.suppressClick = true;
        setTimeout(() => (this.suppressClick = false), 0);
        vp.classList.remove("panning");
      }
      drag = undefined;
    };
    vp.addEventListener("pointerup", end);
    vp.addEventListener("pointercancel", end);
    vp.addEventListener("scroll", () => this.minimap.update());
    // キャプチャで受け取る: ページの指定と、注釈のクリック (要素の処理) の両方を行うため
    vp.addEventListener(
      "click",
      (e) => {
        if (this.suppressClick) {
          e.stopPropagation();
          e.preventDefault();
          return;
        }
        if (this.kind !== "flow") return;
        const page = (e.target as Element).closest?.("g[data-page]");
        if (page) this.selectPageOf(page, false);
        else this.clearSelection();
      },
      true
    );
  }

  // ---- ページの強調と分割画面 (flow) -----------------------------------------------

  private pageElement(column: number, page: number): SVGGElement | undefined {
    return this.viewport.querySelector<SVGGElement>(`[data-column="${column}"] > g[data-page="${page}"]`) ?? undefined;
  }

  private pageCount(column: number): number {
    return this.viewport.querySelectorAll(`[data-column="${column}"] > g[data-page]`).length;
  }

  private columnCount(): number {
    return this.viewport.querySelectorAll("[data-column]").length;
  }

  /** 図の要素を含むページを強調する (ハイライトの一覧から位置を示したときなど) */
  selectPageOf(el: Element, scroll = true): void {
    const g = el.closest("g[data-page]");
    const col = g?.closest<SVGGElement>("[data-column]");
    if (!g || !col) return;
    this.selectPage(Number(col.dataset.column), Number((g as SVGGElement).dataset.page), scroll);
  }

  /** keepSavedPage: 保存時のページは今のページのまま (← → で列だけを変えたとき) */
  selectPage(column: number, page: number, scroll = true, keepSavedPage = false): void {
    if (this.kind !== "flow" || !this.detail) return;
    const g = this.pageElement(column, page);
    if (!g) return;
    for (const el of this.viewport.querySelectorAll(".dra-page-selected")) el.classList.remove("dra-page-selected");
    g.classList.add("dra-page-selected");
    if (this.selected?.column !== column) this.savedPageCount = undefined;
    if (!keepSavedPage) this.savedPage = page;
    this.selected = { column, page };
    const split = clamp(this.hooks.prefs().split ?? 0.6, 0.2, 0.85);
    this.host.style.setProperty("--split", `${split * 100}%`);
    this.host.classList.add("split");
    if (this.detailMode === "saved") this.renderSaved();
    else this.renderDetail(g, column, page);
    if (scroll) g.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  /** 右の画面の表示を切り替える */
  setDetailMode(mode: "schematic" | "saved"): void {
    this.detailMode = mode;
    if (mode === "saved") this.renderSaved();
    else this.refreshDetail();
  }

  /** 保存時のページの状態が変わったとき (作成が終わったなど) */
  refreshSaved(): void {
    if (!this.selected) return;
    if (this.detailMode === "saved") this.renderSaved();
    else this.updateModeButtons(this.selected.column);
  }

  private updateModeButtons(column: number): void {
    if (!this.modeButtons) return;
    const S = this.hooks.strings();
    const info = this.hooks.saved?.info(column);
    this.modeButtons.schematic.setAttribute("aria-pressed", String(this.detailMode === "schematic"));
    this.modeButtons.saved.setAttribute("aria-pressed", String(this.detailMode === "saved"));
    this.modeButtons.saved.title = info ? `${S.detailSaved} (${info.label})` : S.savedNone;
  }

  /** 右の画面に、保存時のページを表示する (無ければ・作成中ならその案内) */
  private renderSaved(): void {
    const body = this.detailBody;
    const sel = this.selected;
    if (!body || !this.detailTitle || !sel || !this.hooks.saved) return;
    const S = this.hooks.strings();
    const column = sel.column;
    const info = this.hooks.saved.info(column);
    this.updateModeButtons(column);
    const title = this.hooks.columnTitle?.(column) ?? "";
    const message = (text: string, action?: { label: string; run: () => void }) => {
      this.detailTitle!.textContent = S.savedPageDetail(title, undefined, undefined, info?.label);
      const p = document.createElement("p");
      p.className = "page-detail-message";
      p.textContent = text;
      if (action) p.append(" ", button(action.label, action.label, action.run));
      body.replaceChildren(p);
    };
    if (!info) return message(S.savedNone);
    if (info.state === "off") return message(S.savedOff);
    if (info.state === "none") {
      this.hooks.saved.request(column);
      return message(S.savedCreating);
    }
    if (info.state === "pending") return message(S.savedCreating);
    if (info.state === "failed") return message(S.savedFailed(info.error ?? ""), { label: S.savedRetry, run: () => this.hooks.saved?.request(column) });
    const cs = getComputedStyle(body);
    const bw = body.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const bh = body.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
    if (bw <= 0 || bh <= 0) return;
    if (info.pageCount !== undefined) this.savedPage = clamp(this.savedPage, 0, Math.max(0, info.pageCount - 1));
    const dpr = window.devicePixelRatio || 1;
    const req = ++this.savedRequest;
    const page = this.savedPage;
    this.detailTitle.textContent = S.savedPageDetail(title, page + 1, info.pageCount, info.label);
    if (this.detailHead) this.detailHead.title = S.savedPageHint;
    void this.hooks.saved
      .render(column, page, bw * dpr, bh * dpr)
      .then((r) => {
        // 古い依頼の結果は捨てる (ページを続けて送ったときなど)
        if (req !== this.savedRequest || this.detailMode !== "saved" || this.selected?.column !== column) return;
        this.savedPageCount = r.pageCount;
        const img = document.createElement("img");
        img.className = "page-detail-image";
        img.alt = S.savedPageDetail(title, page + 1, r.pageCount, info.label);
        img.onload = () => {
          img.style.width = `${img.naturalWidth / dpr}px`;
          img.style.height = `${img.naturalHeight / dpr}px`;
        };
        img.src = r.dataUrl;
        this.detailTitle!.textContent = S.savedPageDetail(title, page + 1, r.pageCount, info.label);
        body.replaceChildren(img);
      })
      .catch((err) => {
        if (req === this.savedRequest) message(S.savedFailed(err instanceof Error ? err.message : String(err)));
      });
  }

  clearSelection(): void {
    if (!this.selected) return;
    for (const el of this.viewport.querySelectorAll(".dra-page-selected")) el.classList.remove("dra-page-selected");
    this.selected = undefined;
    this.host.classList.remove("split");
    this.detailBody?.replaceChildren();
    if (this.detailTitle) this.detailTitle.textContent = "";
  }

  get selectedPage(): { column: number; page: number } | undefined {
    return this.selected;
  }

  /** いまの状態 (スモークテスト用) */
  state(): { zoom: number; left: number; top: number; selected?: { column: number; page: number }; split: boolean; corner?: string } {
    return {
      zoom: this.zoomValue,
      left: this.viewport.scrollLeft,
      top: this.viewport.scrollTop,
      selected: this.selected,
      split: this.host.classList.contains("split"),
      corner: this.minimap.el.dataset.corner,
    };
  }

  /** 強調中のページを描き直す (図の要素の印が変わったとき) */
  refreshDetail(): void {
    if (!this.selected) return;
    if (this.detailMode === "saved") return this.renderSaved();
    const g = this.pageElement(this.selected.column, this.selected.page);
    if (g) this.renderDetail(g, this.selected.column, this.selected.page);
  }

  /** 右の画面: 図を複製し、表示範囲 (viewBox) をページにする */
  private renderDetail(g: SVGGElement, column: number, page: number): void {
    if (!this.svg || !this.detailBody || !this.detailHead) return;
    const frame = g.querySelector<SVGRectElement>("rect.dra-page");
    if (!frame) return;
    const x = frame.x.baseVal.value;
    const y = frame.y.baseVal.value;
    const w = frame.width.baseVal.value;
    const h = frame.height.baseVal.value;
    const clone = this.svg.cloneNode(true) as SVGSVGElement;
    clone.removeAttribute("width");
    clone.removeAttribute("height");
    // ページの縦横比の枠にし (大きさは fitDetail で決める)、ページの外 (隣の帯など) は切り取る
    const m = 4;
    clone.setAttribute("viewBox", `${x - m} ${y - m} ${w + 2 * m} ${h + 2 * m}`);
    this.detailAspect = (w + 2 * m) / (h + 2 * m);
    const NS = "http://www.w3.org/2000/svg";
    const clip = document.createElementNS(NS, "clipPath");
    clip.id = "dra-detail-clip";
    const rect = document.createElementNS(NS, "rect");
    for (const [k, v] of Object.entries({ x: x - 2, y: y - 2, width: w + 4, height: h + 4 })) rect.setAttribute(k, String(v));
    clip.append(rect);
    const content = document.createElementNS(NS, "g");
    content.setAttribute("clip-path", "url(#dra-detail-clip)");
    for (const child of [...clone.childNodes]) if (!(child instanceof SVGDefsElement)) content.append(child);
    clone.append(clip, content);
    clone.setAttribute("preserveAspectRatio", "xMidYMid meet");
    for (const el of clone.querySelectorAll(".dra-page-selected")) el.classList.remove("dra-page-selected");
    const S = this.hooks.strings();
    if (this.detailTitle) this.detailTitle.textContent = S.pageDetail(this.hooks.columnTitle?.(column) ?? "", page + 1);
    this.detailHead.title = S.pageDetailHint;
    this.updateModeButtons(column);
    this.detailBody.replaceChildren(clone);
    this.fitDetail();
    this.hooks.onDetail?.(this.detailBody);
  }

  /** 右の画面のページを、縦横比を保って表示エリアに最大限フィットする大きさにする */
  private fitDetail(): void {
    const body = this.detailBody;
    const svg = body?.querySelector("svg");
    if (!body || !svg) return;
    const cs = getComputedStyle(body);
    const bw = body.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const bh = body.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
    if (bw <= 0 || bh <= 0) return;
    const w = Math.min(bw, bh * this.detailAspect);
    svg.style.width = `${w}px`;
    svg.style.height = `${w / this.detailAspect}px`;
  }

  private initSplitter(): void {
    const sp = this.splitter!;
    let id: number | undefined;
    sp.addEventListener("pointerdown", (e) => {
      id = e.pointerId;
      capture(sp, e.pointerId);
      sp.classList.add("dragging");
    });
    sp.addEventListener("pointermove", (e) => {
      if (id !== e.pointerId) return;
      const r = this.host.getBoundingClientRect();
      const f = clamp((e.clientX - r.left) / r.width, 0.2, 0.85);
      this.host.style.setProperty("--split", `${f * 100}%`);
    });
    const end = (e: PointerEvent) => {
      if (id !== e.pointerId) return;
      id = undefined;
      sp.classList.remove("dragging");
      const r = this.host.getBoundingClientRect();
      const left = this.left.getBoundingClientRect();
      this.hooks.savePrefs({ ...this.hooks.prefs(), split: clamp(left.width / r.width, 0.2, 0.85) });
    };
    sp.addEventListener("pointerup", end);
    sp.addEventListener("pointercancel", end);
  }

  private onKey(e: KeyboardEvent): void {
    if (this.kind !== "flow" || !this.selected || typing(e) || this.host.hidden || !this.visible) return;
    // ミニマップの上にマウスがあるときのカーソルキーは、ミニマップの位置の切り替えに使う
    if (this.minimap.hovered) return;
    const { column, page } = this.selected;
    // 保存時のページを表示しているときの ↑ ↓ は、保存時のページの中で動く
    if (this.detailMode === "saved" && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
      e.preventDefault();
      const count = this.savedPageCount ?? this.hooks.saved?.info(column)?.pageCount ?? 1;
      const nextPage = clamp(this.savedPage + (e.key === "ArrowUp" ? -1 : 1), 0, Math.max(0, count - 1));
      if (nextPage !== this.savedPage) {
        this.savedPage = nextPage;
        this.renderSaved();
      }
      return;
    }
    let next: { column: number; page: number } | undefined;
    switch (e.key) {
      case "ArrowUp":
        if (page > 0) next = { column, page: page - 1 };
        break;
      case "ArrowDown":
        if (page + 1 < this.pageCount(column)) next = { column, page: page + 1 };
        break;
      case "ArrowLeft":
      case "ArrowRight": {
        const c = column + (e.key === "ArrowLeft" ? -1 : 1);
        if (c >= 0 && c < this.columnCount()) next = { column: c, page: Math.min(page, this.pageCount(c) - 1) };
        break;
      }
      case "Escape":
        this.clearSelection();
        return;
      default:
        return;
    }
    e.preventDefault();
    if (next && next.page >= 0) this.selectPage(next.column, next.page, true, this.detailMode === "saved");
  }
}

/** ミニマップ: 図全体の縮小表示と、いまの表示範囲の枠 */
class Minimap {
  readonly el: HTMLDivElement;
  private readonly img: HTMLImageElement;
  private readonly frame: HTMLDivElement;
  private url?: string;
  /** 大きさを変えている間の幅 (終わったら設定に保存する) */
  private dragWidth?: number;
  hovered = false;

  constructor(
    private readonly viewer: FigureViewer,
    private readonly hooks: ViewerHooks
  ) {
    this.el = div("minimap");
    this.el.title = hooks.strings().minimapHint;
    this.img = document.createElement("img");
    this.img.alt = "";
    this.img.draggable = false;
    this.frame = div("minimap-frame");
    this.el.append(this.img, this.frame);
    this.el.hidden = true;
    this.init();
  }

  private prefs(): MinimapPrefs {
    const p = this.hooks.prefs().minimap;
    return { width: clamp(this.dragWidth ?? p?.width ?? 200, MINIMAP_MIN, MINIMAP_MAX), corner: p?.corner ?? "bl" };
  }

  private save(p: MinimapPrefs): void {
    this.hooks.savePrefs({ ...this.hooks.prefs(), minimap: p });
  }

  setImage(svg: string | undefined): void {
    if (this.url) URL.revokeObjectURL(this.url);
    this.url = undefined;
    this.el.hidden = !svg;
    if (!svg) return;
    this.url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
    this.img.src = this.url;
    this.layout();
    this.update();
  }

  /** 大きさ (縦横比は図に合わせる) と位置 (四隅) */
  layout(): void {
    if (this.el.hidden) return;
    const { w: nw, h: nh } = this.viewer.naturalSize;
    const area = this.el.parentElement!.getBoundingClientRect();
    const p = this.prefs();
    let w = p.width;
    let h = (w * nh) / Math.max(1, nw);
    // 表示枠の高さの 60% を超えないようにする
    const maxH = Math.max(60, area.height * 0.6);
    if (h > maxH) {
      h = maxH;
      w = (h * nw) / Math.max(1, nh);
    }
    w = Math.min(w, Math.max(60, area.width * 0.6));
    h = (w * nh) / Math.max(1, nw);
    Object.assign(this.el.style, {
      width: `${w}px`,
      height: `${h}px`,
      left: p.corner[1] === "l" ? `${MINIMAP_MARGIN}px` : "",
      right: p.corner[1] === "r" ? `${MINIMAP_MARGIN + 14}px` : "",
      top: p.corner[0] === "t" ? `${MINIMAP_MARGIN + 34}px` : "",
      bottom: p.corner[0] === "b" ? `${MINIMAP_MARGIN + 14}px` : "",
    });
    this.el.dataset.corner = p.corner;
  }

  /** 表示範囲の枠。図全体が収まっているときは薄く表示する */
  update(): void {
    if (this.el.hidden) return;
    const { w: nw, h: nh } = this.viewer.naturalSize;
    const r = this.viewer.visibleRect();
    const sx = this.el.clientWidth / Math.max(1, nw);
    const sy = this.el.clientHeight / Math.max(1, nh);
    const x = clamp(r.x, 0, nw);
    const y = clamp(r.y, 0, nh);
    Object.assign(this.frame.style, {
      left: `${x * sx}px`,
      top: `${y * sy}px`,
      width: `${Math.max(2, (Math.min(r.x + r.w, nw) - x) * sx)}px`,
      height: `${Math.max(2, (Math.min(r.y + r.h, nh) - y) * sy)}px`,
    });
    this.el.classList.toggle("faded", this.viewer.fitsEntirely());
  }

  /** 大きさを変える辺 (ミニマップの位置に応じて内側の辺) */
  private edgeAt(e: PointerEvent): { h?: "left" | "right"; v?: "top" | "bottom" } {
    const r = this.el.getBoundingClientRect();
    const corner = this.prefs().corner;
    const out: { h?: "left" | "right"; v?: "top" | "bottom" } = {};
    if (corner[1] === "l" && r.right - e.clientX <= EDGE) out.h = "right";
    if (corner[1] === "r" && e.clientX - r.left <= EDGE) out.h = "left";
    if (corner[0] === "b" && e.clientY - r.top <= EDGE) out.v = "top";
    if (corner[0] === "t" && r.bottom - e.clientY <= EDGE) out.v = "bottom";
    return out;
  }

  private moveTo(e: PointerEvent | MouseEvent): void {
    const r = this.el.getBoundingClientRect();
    const { w: nw, h: nh } = this.viewer.naturalSize;
    this.viewer.centerOn(((e.clientX - r.left) / r.width) * nw, ((e.clientY - r.top) / r.height) * nh);
  }

  private init(): void {
    const el = this.el;
    let mode: { kind: "move" } | { kind: "resize"; edge: ReturnType<Minimap["edgeAt"]>; x: number; y: number; w: number; h: number } | undefined;
    el.addEventListener("pointerdown", (e) => {
      if (e.button !== 0) return;
      e.stopPropagation();
      e.preventDefault();
      capture(el, e.pointerId);
      const edge = this.edgeAt(e);
      if (edge.h || edge.v) {
        mode = { kind: "resize", edge, x: e.clientX, y: e.clientY, w: el.clientWidth, h: el.clientHeight };
      } else {
        mode = { kind: "move" };
        this.moveTo(e);
      }
    });
    el.addEventListener("pointermove", (e) => {
      if (!mode) {
        const edge = this.edgeAt(e);
        el.style.cursor = edge.h && edge.v ? (edge.h === "right") === (edge.v === "bottom") ? "nwse-resize" : "nesw-resize" : edge.h ? "ew-resize" : edge.v ? "ns-resize" : "";
        return;
      }
      if (mode.kind === "move") {
        this.moveTo(e);
        return;
      }
      const { w: nw, h: nh } = this.viewer.naturalSize;
      const aspect = nh / Math.max(1, nw);
      const dw = mode.edge.h ? (mode.edge.h === "right" ? e.clientX - mode.x : mode.x - e.clientX) : 0;
      const dh = mode.edge.v ? (mode.edge.v === "bottom" ? e.clientY - mode.y : mode.y - e.clientY) : 0;
      const w = Math.max(mode.w + dw, mode.edge.v ? (mode.h + dh) / aspect : 0);
      this.dragWidth = clamp(w, MINIMAP_MIN, MINIMAP_MAX);
      this.layout();
      this.update();
    });
    const end = (e: PointerEvent) => {
      if (mode?.kind === "resize") {
        const p = this.prefs();
        this.dragWidth = undefined;
        this.save(p);
      }
      mode = undefined;
      if (el.hasPointerCapture?.(e.pointerId)) el.releasePointerCapture(e.pointerId);
    };
    el.addEventListener("pointerup", end);
    el.addEventListener("pointercancel", end);
    el.addEventListener("click", (e) => e.stopPropagation());
    el.addEventListener("dblclick", (e) => {
      e.stopPropagation();
      this.viewer.fit();
    });
    el.addEventListener("mouseenter", () => (this.hovered = true));
    el.addEventListener("mouseleave", () => (this.hovered = false));
    el.addEventListener("wheel", (e) => e.stopPropagation());
    document.addEventListener("keydown", (e) => {
      if (!this.hovered || el.hidden || typing(e)) return;
      const p = this.prefs();
      let [v, h] = [p.corner[0], p.corner[1]];
      if (e.key === "ArrowLeft") h = "l";
      else if (e.key === "ArrowRight") h = "r";
      else if (e.key === "ArrowUp") v = "t";
      else if (e.key === "ArrowDown") v = "b";
      else return;
      e.preventDefault();
      this.save({ ...p, corner: `${v}${h}` as Corner });
      this.layout();
      this.update();
    });
  }
}
