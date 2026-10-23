/**
 * flowSvg.ts
 *
 * buildFlow の結果を1枚の SVG に描画する。
 *  - 文書の模式的なページを縦に並べた列 (サムネイル) を、左から右へ時系列順に並べる。
 *    先頭の列は最初の区間の開始時点、以降の列は各区間の終了時点の文書。
 *    区間と区間の間には変更が無いため、前の区間の終了時点と次の区間の開始時点は同じ文書であり、
 *    1つの列を共有する (開始時点の列を繰り返さない)。
 *  - 各区間の終了時点の列では、区間内に編集された段落を、細かい編集は緑、一括挿入はオレンジで塗る。
 *    次の区間で削除される段落・図表には、ページ右端に赤い印を付ける。
 *  - 列と列の間には、段落・図表 (表は1つの単位) ごとに、区間の開始時点と終了時点の位置・高さを
 *    帯で結んだスロープ図を描く (帯の太さの変化で増減が、行き先の無い帯で削除が分かる)。
 *    移動・並べ替えは、移動元の位置から移動先の位置へ青い帯で結ぶ (順序が入れ替われば他の帯と交差する)。
 *  - 各区間の帯の下に、区間の開始・終了日時と、のべ追加・削除文字数のキャプションを付け、
 *    列の見出しに、その後の無編集期間の長さを表記する。
 */

import { FlowResult, FlowSession, PageLayout, ParaHeat, SlopeUnit } from "./flow";
import type { PageItem } from "./pageLayout";
import { esc, formatGapHours, renderNote, TitleLink, titleMarkup } from "./svgChart";
import { estimateLabelWidth, t } from "./i18n";
import { Category, CategoryRegistry, patternOverlay, renderPatternDefs } from "./categories";
import {
  decorateElement,
  DecorateOptions,
  flowBandKey,
  flowCaptionKey,
  flowMoveKey,
  flowParaKey,
  ResolvedAnnotation,
} from "./figureTargets";

export interface FlowSvgOptions {
  title?: string;
  /** ページのサムネイルの幅 (px) */
  pageWidth?: number;
  /** 開始時点と終了時点のサムネイルの間 (スロープ図) の幅 (px) */
  slopeWidth?: number;
  fontFamily?: string;
  /** 図の右下に小さく添える注記 (判定ルールの識別子など) */
  note?: string;
  /** 色のカテゴリ (省略時は既定のカテゴリ) */
  categories?: CategoryRegistry;
  /**
   * 列を <g data-column="k"> で囲み (k = 0 は区間1の開始時点、k は区間 k の終了時点)、段落の要素に
   * data-para (段落の番号) を付ける (アプリで該当箇所を探すため)
   */
  annotate?: boolean;
  /**
   * 部分ごとの注釈 (マウスオーバーの説明・クリックで開くリンク)。figureTargets.ts の resolveAnnotations で作る。
   * 部分のキーは段落 flow:para:<列>:<段落>、帯 flow:band:<区間>:<単位>、移動の帯、キャプション
   */
  annotations?: Map<string, ResolvedAnnotation>;
  /** 見出しの一部 (ファイル名など) をリンクにする (OneDrive / SharePoint の元の文書の URL) */
  titleLink?: TitleLink;
}

/** これ未満の度合いは塗らない */
const MIN_INTENSITY = 0.05;

const PAGE_GAP = 10;
const MARGIN = 24;
const HEADER_HEIGHT = 128;
/** 「開始時点」「終了時点」の見出しの高さ */
const STACK_LABEL_HEIGHT = 34;
const CAPTION_HEIGHT = 78;

function pad(n: number): string {
  return String(n).padStart(2, "0");
}

function fmtDateTime(d: Date): string {
  return `${pad(d.getMonth() + 1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** 開始・終了が同じ日なら、終了側は時刻だけにする */
function fmtRange(start: Date, end: Date): string {
  const sameDay = start.toDateString() === end.toDateString();
  return `${fmtDateTime(start)} – ${sameDay ? `${pad(end.getHours())}:${pad(end.getMinutes())}` : fmtDateTime(end)}`;
}

function fmtCount(n: number): string {
  return n.toLocaleString("en-US");
}

function fill(rgb: string, intensity: number): string {
  return `rgba(${rgb},${(0.15 + 0.75 * intensity).toFixed(3)})`;
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** 段落の塗り方: 主の色で段落全体、副の色を左余白の細い帯で。deletedNext なら右余白に赤い印 */
interface Tint {
  main?: string;
  mainV?: number;
  /** 主の色のカテゴリ (模様を重ねるため) */
  mainCat?: Category;
  sub?: string;
  subV?: number;
  /** 次の区間で削除される (赤) / 移動していく (青) ことを示す右余白の印の色 */
  nextMark?: string;
}

/** 段落の行・図 (浮動配置の図は除く) を、つながった矩形ごとにまとめる (段組みでは段ごとの矩形になる) */
function paraBoxes(page: PageLayout): Map<number, Box[]> {
  const out = new Map<number, Box[]>();
  const TOL = 30;
  for (const it of page.items) {
    if (!("paraIndex" in it) || (it.type === "figure" && it.float)) continue;
    const list = out.get(it.paraIndex) ?? [];
    out.set(it.paraIndex, list);
    const last = list[list.length - 1];
    const touches =
      last &&
      it.x < last.x + last.w + TOL &&
      it.x + it.w > last.x - TOL &&
      it.y < last.y + last.h + TOL &&
      it.y + it.h > last.y - TOL;
    if (!touches) {
      list.push({ x: it.x, y: it.y, w: it.w, h: it.h });
      continue;
    }
    const x = Math.min(last.x, it.x);
    const y = Math.min(last.y, it.y);
    last.w = Math.max(last.x + last.w, it.x + it.w) - x;
    last.h = Math.max(last.y + last.h, it.y + it.h) - y;
    last.x = x;
    last.y = y;
  }
  return out;
}

/** 行の棒 (行の長さ fill を、配置に合わせて置く)。縦書きのページでは縦の棒 */
function lineBar(it: { x: number; y: number; w: number; h: number; fill: number }, vertical: boolean, thick: number, align = "left"): Box {
  const along = vertical ? it.h : it.w;
  const across = vertical ? it.w : it.h;
  const len = Math.max(1, along * it.fill);
  const off = align === "center" ? (along - len) / 2 : align === "right" ? along - len : 0;
  const th = Math.max(0.6, across * thick);
  return vertical
    ? { x: it.x + (across - th) / 2, y: it.y + off, w: th, h: len }
    : { x: it.x + off, y: it.y + (across - th) / 2, w: len, h: th };
}

function renderPage(
  page: PageLayout,
  tintOf: (paraIndex: number) => Tint | undefined,
  x: number,
  y: number,
  k: number,
  deco: DecorateOptions & { column: number } = { column: 0 }
): string {
  const parts: string[] = [];
  const annotate = !!deco.annotate;
  const dataPara = (i: number) => (annotate ? ` data-para="${i}"` : "");
  // 注釈を付けるときは、段落全体を覆う透明な矩形を最後に重ね、マウス操作を受けるようにする (細い行の上でなくても反応する)
  const decorated = annotate || !!deco.annotations?.size;
  const hits: string[] = [];
  const f = (n: number) => n.toFixed(1);
  const px = (b: Box) => ({ x: x + b.x * k, y: y + b.y * k, w: b.w * k, h: b.h * k });
  const vertical = !!page.vertical;
  parts.push(`<rect x="${x}" y="${y}" width="${f(page.width * k)}" height="${f(page.height * k)}" fill="#ffffff" stroke="#b8b8b8"/>`);

  // 表のセルの枠・段の境界線
  for (const it of page.items) {
    if (it.type !== "cell") continue;
    const b = px(it);
    parts.push(`<rect x="${f(b.x)}" y="${f(b.y)}" width="${f(b.w)}" height="${f(b.h)}" fill="none" stroke="#c9ced6" stroke-width="0.5"/>`);
  }
  for (const r of page.rules ?? []) {
    parts.push(
      `<line x1="${f(x + r.x1 * k)}" y1="${f(y + r.y1 * k)}" x2="${f(x + r.x2 * k)}" y2="${f(y + r.y2 * k)}" stroke="#b8b8b8" stroke-width="0.5"/>`
    );
  }

  // 段落ごとの塗り (このページに載っている範囲)
  for (const [paraIndex, boxes] of paraBoxes(page)) {
    const t = tintOf(paraIndex);
    for (const box of boxes) {
      const b = px(box);
      const rx = f(b.x - 2);
      const rw = f(b.w + 4);
      if (decorated) {
        hits.push(
          decorateElement(
            `<rect x="${rx}" y="${f(b.y)}" width="${rw}" height="${f(b.h)}" fill="transparent"/>`,
            flowParaKey(deco.column, paraIndex),
            deco
          )
        );
      }
      if (!t) continue;
      if (t.main && t.mainV) {
        parts.push(
          `<rect x="${rx}" y="${f(b.y)}" width="${rw}" height="${f(b.h)}" fill="${fill(t.main, t.mainV)}"${dataPara(paraIndex)}/>` +
            (t.mainCat ? patternOverlay(t.mainCat, rx, f(b.y), rw, f(b.h)) : "")
        );
      }
      if (t.sub && t.subV) {
        parts.push(`<rect x="${f(b.x - 7)}" y="${f(b.y)}" width="4" height="${f(b.h)}" fill="${fill(t.sub, t.subV)}"/>`);
      }
      if (t.nextMark) {
        parts.push(`<rect x="${f(b.x + b.w + 3)}" y="${f(b.y)}" width="4" height="${f(b.h)}" fill="rgb(${t.nextMark})"/>`);
      }
    }
  }

  // 本文の行・図・脚注
  for (const it of page.items as PageItem[]) {
    if (it.type === "cell") continue;
    if (it.type === "figure") {
      const b = px(it);
      const fx = b.x + b.w * 0.02;
      const fy = b.y + b.h * 0.02;
      const fw = b.w * 0.96;
      const fh = b.h * 0.96;
      const dash = it.float ? ` stroke-dasharray="2,1.5"` : "";
      parts.push(
        `<rect x="${f(fx)}" y="${f(fy)}" width="${f(fw)}" height="${f(fh)}" fill="#eceff3" stroke="#9aa3ad" stroke-width="0.6"${dash}/>`,
        `<path d="M${f(fx)},${f(fy)} L${f(fx + fw)},${f(fy + fh)} M${f(fx + fw)},${f(fy)} L${f(fx)},${f(fy + fh)}" stroke="#c4cad1" stroke-width="0.6"/>`
      );
      continue;
    }
    if (it.fill <= 0) continue;
    if (it.type === "note" || it.type === "repeat") {
      const b = px(lineBar(it, vertical, 0.4));
      parts.push(`<rect x="${f(b.x)}" y="${f(b.y)}" width="${f(Math.max(1, b.w))}" height="${f(Math.max(0.6, b.h))}" fill="#c3c7cd"/>`);
      continue;
    }
    const b = px(lineBar(it, vertical, it.style === "heading" ? 0.6 : 0.45, it.align));
    const color = it.style === "heading" ? "#4a4f57" : it.style === "table" ? "#8ea0b8" : "#a3a8b0";
    parts.push(
      `<rect x="${f(b.x)}" y="${f(b.y)}" width="${f(Math.max(vertical ? 0.6 : 1, b.w))}" height="${f(Math.max(vertical ? 1 : 0.6, b.h))}" fill="${color}"${dataPara(it.paraIndex)}/>`
    );
  }
  parts.push(...hits);
  return parts.join("\n");
}

/** ページ列の高さ (px) */
function stackHeight(pages: PageLayout[], k: number): number {
  return pages.reduce((n, p) => n + p.height * k, 0) + Math.max(0, pages.length - 1) * PAGE_GAP;
}

/**
 * 縦に積んだページ列の中での、段落ごとの縦の範囲 (px)。帯の端に使う。
 * 段組みで段をまたぐ段落は、矩形の高さを足した長さにする (ページの高さいっぱいに広がらないように)。
 * 縦書きのページでは、右端からの位置をページの高さに写す
 */
function stackPositions(pages: PageLayout[], top: number, k: number): Map<number, { top: number; bottom: number }> {
  const spans = new Map<number, { top: number; bottom: number }[]>();
  let pageTop = top;
  for (const page of pages) {
    for (const [i, boxes] of paraBoxes(page)) {
      const list = spans.get(i) ?? [];
      spans.set(i, list);
      for (const b of boxes) {
        const [t, h] = page.vertical
          ? [((page.width - (b.x + b.w)) / page.width) * page.height, (b.w / page.width) * page.height]
          : [b.y, b.h];
        list.push({ top: pageTop + t * k, bottom: pageTop + (t + h) * k });
      }
    }
    pageTop += page.height * k + PAGE_GAP;
  }
  const pos = new Map<number, { top: number; bottom: number }>();
  for (const [i, list] of spans) {
    if (list.length === 0) continue;
    const first = list[0];
    const last = list[list.length - 1];
    const bottom =
      last.top >= first.bottom - 0.5 ? last.bottom : first.top + list.reduce((n, s) => n + (s.bottom - s.top), 0);
    pos.set(i, { top: first.top, bottom: Math.max(bottom, first.bottom) });
  }
  return pos;
}

function unitRange(
  unit: SlopeUnit,
  pos: Map<number, { top: number; bottom: number }>
): { top: number; bottom: number } | undefined {
  let range: { top: number; bottom: number } | undefined;
  for (const i of unit.paraIndexes) {
    const r = pos.get(i);
    if (!r) continue;
    range = range ? { top: Math.min(range.top, r.top), bottom: Math.max(range.bottom, r.bottom) } : { ...r };
  }
  return range;
}

function unitColor(u: SlopeUnit, reg: CategoryRegistry): string {
  switch (u.change) {
    case "deleted":
      return reg.rgb(reg.byRole("deleted"));
    case "added":
      return reg.rgb(u.addedByBulk ? reg.highlightOrDefault(u.category) : reg.byRole("fine"));
    case "bulk":
      return reg.rgb(reg.highlightOrDefault(u.category));
    case "fine":
      return reg.rgb(reg.byRole("fine"));
    case "movedFrom":
    case "movedTo":
    case "moved":
      return reg.rgb(reg.byRole("moved"));
    default:
      return reg.rgb(reg.byRole("unchanged"));
  }
}

/**
 * 開始時点 (lx) と終了時点 (rx) の間に、単位ごとの帯を描く。
 * 両方にある単位は台形の帯、削除された単位は中央へすぼむ帯、追加された単位は中央から広がる帯。
 * 移動・並べ替えで対応づいた単位は、移動元から移動先へ青い帯で結び、最後に重ねる。
 */
function renderSlopes(
  s: FlowSession,
  lx: number,
  rx: number,
  stackTop: number,
  k: number,
  reg: CategoryRegistry,
  sessionIndex = 0,
  deco: DecorateOptions = {}
): string {
  const BLUE = reg.rgb(reg.byRole("moved"));
  const left = stackPositions(s.startPages, stackTop, k);
  const right = stackPositions(s.endPages, stackTop, k);
  const mx = (lx + rx) / 2;
  const f = (n: number) => n.toFixed(1);
  // 帯の端を滑らかにつなぐ3次ベジェ曲線 (x0,y0) → (x1,y1)
  const curve = (x0: number, y0: number, x1: number, y1: number) =>
    `C${f((x0 + x1) / 2)},${f(y0)} ${f((x0 + x1) / 2)},${f(y1)} ${f(x1)},${f(y1)}`;

  const band = (L: { top: number; bottom: number }, R: { top: number; bottom: number }) =>
    `M${f(lx)},${f(L.top)} ${curve(lx, L.top, rx, R.top)} ` +
    `L${f(rx)},${f(R.bottom)} ${curve(rx, R.bottom, lx, L.bottom)} Z`;

  // 移動元・移動先として対応づいた単位は、すぼむ/広がる帯の代わりに移動の帯で結ぶ
  const linkedFrom = new Set(s.links.map((l) => l.fromKey));
  const linkedTo = new Set(s.links.map((l) => l.toKey));

  // 変化の無い帯を先に描き、変化のあった帯を上に重ねる
  const ordered = [...s.units].sort((a, b) => Number(a.change !== "unchanged") - Number(b.change !== "unchanged"));
  const parts: string[] = [];
  for (const u of ordered) {
    if ((u.change === "movedFrom" && linkedFrom.has(u.key)) || (u.change === "movedTo" && linkedTo.has(u.key))) {
      continue;
    }
    const L = unitRange(u, left);
    const R = unitRange(u, right);
    const color = unitColor(u, reg);
    const opacity = u.change === "unchanged" ? 0.22 : 0.3 + 0.5 * u.intensity;
    let d: string | undefined;
    if (L && R) {
      d = band(L, R);
    } else if (L) {
      const c = (L.top + L.bottom) / 2;
      d = `M${f(lx)},${f(L.top)} ${curve(lx, L.top, mx, c)} ${curve(mx, c, lx, L.bottom)} Z`;
    } else if (R) {
      const c = (R.top + R.bottom) / 2;
      d = `M${f(mx)},${f(c)} ${curve(mx, c, rx, R.top)} L${f(rx)},${f(R.bottom)} ${curve(rx, R.bottom, mx, c)} Z`;
    }
    if (!d) continue;
    parts.push(
      decorateElement(
        `<path d="${d}" fill="rgb(${color})" fill-opacity="${opacity.toFixed(2)}" stroke="rgb(${color})" stroke-opacity="${Math.min(1, opacity + 0.2).toFixed(2)}" stroke-width="0.5"/>`,
        flowBandKey(sessionIndex, u.key),
        deco
      )
    );
  }

  // 移動の帯: 移動元の単位の開始時点の位置 → 移動先の単位の終了時点の位置
  const unitByKey = new Map(s.units.map((u) => [u.key, u]));
  for (const l of s.links) {
    const from = unitByKey.get(l.fromKey);
    const to = unitByKey.get(l.toKey);
    const L = from && unitRange(from, left);
    const R = to && unitRange(to, right);
    if (!L || !R) continue;
    // 段落ごと移動した場合は段落の高さの帯、段落の一部だけが移動した場合は細い帯
    const whole = from!.change === "movedFrom" && to!.change === "movedTo";
    const narrow = (r: { top: number; bottom: number }) => {
      const c = (r.top + r.bottom) / 2;
      return { top: c - 1.5, bottom: c + 1.5 };
    };
    parts.push(
      decorateElement(
        `<path d="${band(whole ? L : narrow(L), whole ? R : narrow(R))}" fill="rgb(${BLUE})" fill-opacity="0.55" stroke="rgb(${BLUE})" stroke-opacity="0.9" stroke-width="0.6"/>`,
        flowMoveKey(sessionIndex, l.fromKey, l.toKey),
        deco
      )
    );
  }
  return parts.join("\n");
}

/** 区間 index (0始まり) のキャプション。x〜x+width は前後の列の中心の間 */
function renderCaption(
  s: FlowSession,
  index: number,
  x: number,
  y: number,
  width: number,
  font: string,
  reg: CategoryRegistry,
  deco: DecorateOptions = {}
): string {
  const RED = reg.rgb(reg.byRole("deleted"));
  // 一括挿入の色は、区間で最も多くを占めたハイライトのカテゴリ
  const topCat = Object.entries(s.highlightChars).reduce<[string, number] | undefined>(
    (a, e) => (!a || e[1] > a[1] ? e : a),
    undefined
  )?.[0];
  const cx = x + width / 2;
  const lines = [
    // どの帯の区間かが分かるよう、前後の列の中心を結ぶ括弧を描く
    `<path d="M${x},${y - 6} L${x},${y} L${x + width},${y} L${x + width},${y - 6}" fill="none" stroke="#bbb"/>`,
    `<text x="${cx}" y="${y + 16}" text-anchor="middle" font-size="12" fill="#333">${esc(
      t("flowSession", index + 1, fmtRange(s.start, s.end))
    )}</text>`,
    `<text x="${cx}" y="${y + 33}" text-anchor="middle" font-size="12">` +
      `<tspan fill="#1a7f37">${esc(t("flowChars", "+", fmtCount(s.insChars)))}</tspan>` +
      `<tspan fill="#888"> / </tspan>` +
      `<tspan fill="rgb(${RED})">${esc(t("flowChars", "−", fmtCount(s.delChars)))}</tspan></text>`,
  ];
  // 一括挿入・移動は1行ずつ (英語では長くなり、隣の区間のキャプションと重なるため)
  const notes: [string, string][] = [];
  if (s.bulkInsChars > 0) notes.push([reg.rgb(reg.highlightOrDefault(topCat)), t("flowBulkNote", fmtCount(s.bulkInsChars))]);
  if (s.movedChars > 0) notes.push([reg.rgb(reg.byRole("moved")), t("flowMovedNote", fmtCount(s.movedChars))]);
  notes.forEach(([rgb, text], i) => {
    lines.push(
      `<text x="${cx}" y="${y + 49 + i * 15}" text-anchor="middle" font-size="11" fill="rgb(${rgb})">${esc(text)}</text>`
    );
  });
  return decorateElement(`<g font-family="${font}">${lines.join("")}</g>`, flowCaptionKey(index), deco);
}

/** 凡例。返り値の width は凡例全体の幅 (図の幅を決めるのに使う) */
function renderLegend(x: number, y: number, font: string, reg: CategoryRegistry): { svg: string; width: number } {
  const rows: Category[][] = [
    [reg.byRole("fine"), ...reg.highlights()],
    [reg.byRole("deleted"), reg.byRole("moved"), reg.byRole("unchanged")],
  ];
  const parts: string[] = [];
  let width = 0;
  rows.forEach((row, r) => {
    let dx = 0;
    const dy = r * 20;
    for (const c of row) {
      const rgb = reg.rgb(c);
      const label = reg.flowLegend(c);
      parts.push(
        `<rect x="${x + dx}" y="${y + dy - 10}" width="12" height="12" fill="${fill(rgb, 0.3)}"/>` +
          `<rect x="${x + dx + 12}" y="${y + dy - 10}" width="12" height="12" fill="${fill(rgb, 1)}"/>` +
          patternOverlay(c, x + dx, y + dy - 10, 24, 12) +
          `<text x="${x + dx + 30}" y="${y + dy}" font-size="12" fill="#333">${esc(label)}</text>`
      );
      dx += 30 + estimateLabelWidth(label, 12) + 24;
    }
    width = Math.max(width, dx);
  });
  const notes = [t("flowLegendNote1"), t("flowLegendNote2")];
  notes.forEach((note, i) => {
    parts.push(`<text x="${x}" y="${y + 40 + i * 15}" font-size="11" fill="#777">${esc(note)}</text>`);
    width = Math.max(width, estimateLabelWidth(note, 11));
  });
  return { svg: `<g font-family="${font}">${parts.join("")}</g>`, width };
}

export function renderFlowSvg(result: FlowResult, opts: FlowSvgOptions = {}): string {
  const sessions = result.sessions;
  if (sessions.length === 0) {
    throw new Error(t("errNoFlowSessions"));
  }
  const font =
    opts.fontFamily ??
    "'Noto Sans CJK JP', 'Noto Sans JP', 'Yu Gothic', 'Hiragino Sans', Meiryo, 'Helvetica Neue', Arial, sans-serif";
  const pageWidth = opts.pageWidth ?? 150;
  const slopeWidth = opts.slopeWidth ?? 72;

  // 列: [区間1の開始時点, 区間1の終了時点, 区間2の終了時点, ...]
  const columns: PageLayout[][] = [sessions[0].startPages, ...sessions.map((s) => s.endPages)];
  // 縮尺 (px / twip): 最も幅の広いページを pageWidth に合わせる (横向きのページが混ざっても列からはみ出さない)
  const maxPageWidth = Math.max(1, ...columns.flatMap((c) => c.map((p) => p.width)));
  const scale = pageWidth / maxPageWidth;
  const columnX = (k: number) => MARGIN + k * (pageWidth + slopeWidth);
  const stackTop = HEADER_HEIGHT + STACK_LABEL_HEIGHT;
  const columnsHeight = Math.max(...columns.map((c) => stackHeight(c, scale)));
  const captionY = stackTop + columnsHeight + 20;
  const reg = opts.categories ?? new CategoryRegistry();
  const legend = renderLegend(MARGIN, 58, font, reg);
  const width = Math.ceil(
    Math.max(720, MARGIN * 2 + legend.width, MARGIN * 2 + columns.length * pageWidth + sessions.length * slopeWidth)
  );
  const height = captionY + CAPTION_HEIGHT + MARGIN;

  const parts: string[] = [];
  parts.push(`<rect width="100%" height="100%" fill="#ffffff"/>`);
  const defs = renderPatternDefs(reg.patterns());
  if (defs) parts.push(defs);
  parts.push(
    `<text x="${MARGIN}" y="32" font-family="${font}" font-size="18" font-weight="bold" fill="#222">${titleMarkup(opts.title ?? t("flowDefaultTitle"), opts.titleLink)}</text>`
  );
  parts.push(legend.svg);

  columns.forEach((pages, k) => {
    const x = columnX(k);
    // この列を終了時点とする区間 (先頭の列には無い) と、この列を開始時点とする区間 (末尾の列には無い)
    const ended = k > 0 ? sessions[k - 1] : undefined;
    const next = sessions[k];

    // 見出し: どの時点の文書か、その後の無編集期間
    const title = ended ? t("flowColumnEnd", k) : t("flowColumnStart");
    parts.push(
      `<text x="${x + pageWidth / 2}" y="${HEADER_HEIGHT + 12}" text-anchor="middle" font-family="${font}" font-size="11" fill="#555">${esc(title)}</text>`
    );
    if (ended && next && next.gapBeforeHours !== null) {
      parts.push(
        `<text x="${x + pageWidth / 2}" y="${HEADER_HEIGHT + 26}" text-anchor="middle" font-family="${font}" font-size="10" fill="#999">${esc(
          t("flowIdle", formatGapHours(next.gapBeforeHours))
        )}</text>`
      );
    }

    const nextMark = new Map<number, string>();
    for (const u of next?.units ?? []) {
      if (u.change !== "deleted" && u.change !== "movedFrom") continue;
      for (const i of u.paraIndexes) {
        nextMark.set(i, reg.rgb(reg.byRole(u.change === "deleted" ? "deleted" : "moved")));
      }
    }
    const tintOf = (i: number): Tint | undefined => {
      const t: Tint = { nextMark: nextMark.get(i) };
      const h: ParaHeat | undefined = ended?.heat.get(i);
      if (h) {
        // priority の高いカテゴリ (同じなら度合いの大きいもの) で段落を塗り、2番目を左余白の細い帯で示す。
        // ハイライトの度合いは、一括挿入の度合いをカテゴリの文字数で按分したもの
        const ranked = [
          { c: reg.byRole("fine"), v: h.green },
          ...Object.entries(h.highlightChars).map(([id, n]) => ({
            c: reg.highlightOrDefault(id),
            v: h.orange * (n / h.bulkInsChars),
          })),
          { c: reg.byRole("moved"), v: h.moved },
        ]
          .filter((e) => e.v >= MIN_INTENSITY)
          .sort((a, b) => b.c.priority - a.c.priority || b.v - a.v);
        if (ranked[0]) Object.assign(t, { main: reg.rgb(ranked[0].c), mainV: ranked[0].v, mainCat: ranked[0].c });
        if (ranked[1]) Object.assign(t, { sub: reg.rgb(ranked[1].c), subV: ranked[1].v });
      }
      return t.main || t.nextMark ? t : undefined;
    };
    if (opts.annotate) parts.push(`<g data-column="${k}">`);
    let pageTop = stackTop;
    for (const page of pages) {
      parts.push(renderPage(page, tintOf, x, pageTop, scale, { ...opts, column: k }));
      pageTop += page.height * scale + PAGE_GAP;
    }
    if (opts.annotate) parts.push(`</g>`);
  });

  sessions.forEach((s, i) => {
    parts.push(renderSlopes(s, columnX(i) + pageWidth, columnX(i + 1), stackTop, scale, reg, i, opts));
    parts.push(
      renderCaption(s, i, columnX(i) + pageWidth / 2, captionY, pageWidth + slopeWidth, font, reg, opts)
    );
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height.toFixed(0)}" viewBox="0 0 ${width} ${height.toFixed(0)}">
${parts.join("\n")}${renderNote(opts.note, width, Math.round(height), ` font-family="${font}"`)}
</svg>
`;
}
