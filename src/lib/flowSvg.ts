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
import { esc, formatGapHours, renderNote } from "./svgChart";
import { estimateLabelWidth, t } from "./i18n";

export interface FlowSvgOptions {
  title?: string;
  /** ページのサムネイルの幅 (px) */
  pageWidth?: number;
  /** 開始時点と終了時点のサムネイルの間 (スロープ図) の幅 (px) */
  slopeWidth?: number;
  fontFamily?: string;
  /** 図の右下に小さく添える注記 (判定ルールの識別子など) */
  note?: string;
}

const GREEN = "22,128,61";
const ORANGE = "234,108,0";
const RED = "198,40,40";
const GRAY = "150,155,165";
const BLUE = "37,99,235";
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

interface PageGeom {
  width: number;
  height: number;
  textX: number;
  textY: number;
  textWidth: number;
  lineHeight: number;
}

/** 段落の塗り方: 主の色で段落全体、副の色を左余白の細い帯で。deletedNext なら右余白に赤い印 */
interface Tint {
  main?: string;
  mainV?: number;
  sub?: string;
  subV?: number;
  /** 次の区間で削除される (赤) / 移動していく (青) ことを示す右余白の印の色 */
  nextMark?: string;
}

function renderPage(
  page: PageLayout,
  tintOf: (paraIndex: number) => Tint | undefined,
  x: number,
  y: number,
  g: PageGeom
): string {
  const parts: string[] = [];
  parts.push(
    `<rect x="${x}" y="${y}" width="${g.width}" height="${g.height.toFixed(1)}" fill="#ffffff" stroke="#b8b8b8"/>`
  );

  // 段落ごとの塗り (このページに載っている範囲)
  for (const [paraIndex, r] of paraRanges([page])) {
    const t = tintOf(paraIndex);
    if (!t) continue;
    const top = y + g.textY + r.top * g.lineHeight;
    const height = (r.bottom - r.top) * g.lineHeight;
    if (t.main && t.mainV) {
      parts.push(
        `<rect x="${(x + g.textX - 2).toFixed(1)}" y="${top.toFixed(1)}" width="${(g.textWidth + 4).toFixed(1)}" height="${height.toFixed(1)}" fill="${fill(t.main, t.mainV)}"/>`
      );
    }
    if (t.sub && t.subV) {
      parts.push(
        `<rect x="${(x + g.textX - 7).toFixed(1)}" y="${top.toFixed(1)}" width="4" height="${height.toFixed(1)}" fill="${fill(t.sub, t.subV)}"/>`
      );
    }
    if (t.nextMark) {
      parts.push(
        `<rect x="${(x + g.textX + g.textWidth + 3).toFixed(1)}" y="${top.toFixed(1)}" width="4" height="${height.toFixed(1)}" fill="rgb(${t.nextMark})"/>`
      );
    }
  }

  // 本文の行・図
  for (const it of page.items) {
    const top = y + g.textY + it.y * g.lineHeight;
    const h = it.h * g.lineHeight;
    if (it.type === "figure") {
      const fx = x + g.textX + g.textWidth * 0.15;
      const fw = g.textWidth * 0.7;
      const fy = top + h * 0.05;
      const fh = h * 0.9;
      parts.push(
        `<rect x="${fx.toFixed(1)}" y="${fy.toFixed(1)}" width="${fw.toFixed(1)}" height="${fh.toFixed(1)}" fill="#eceff3" stroke="#9aa3ad" stroke-width="0.6"/>`,
        `<path d="M${fx.toFixed(1)},${fy.toFixed(1)} L${(fx + fw).toFixed(1)},${(fy + fh).toFixed(1)} M${(fx + fw).toFixed(1)},${fy.toFixed(1)} L${fx.toFixed(1)},${(fy + fh).toFixed(1)}" stroke="#c4cad1" stroke-width="0.6"/>`
      );
      continue;
    }
    if (it.widthFrac <= 0) continue;
    const barH = it.style === "heading" ? h * 0.6 : h * 0.45;
    const color = it.style === "heading" ? "#4a4f57" : it.style === "table" ? "#8ea0b8" : "#a3a8b0";
    parts.push(
      `<rect x="${(x + g.textX).toFixed(1)}" y="${(top + (h - barH) / 2).toFixed(1)}" width="${Math.max(1, g.textWidth * it.widthFrac).toFixed(1)}" height="${Math.max(0.6, barH).toFixed(2)}" fill="${color}"/>`
    );
  }
  return parts.join("\n");
}

/** 段落ごとの縦の範囲 (行単位)。pages を縦に積んだ通しの位置ではなく、各ページ内の位置 */
function paraRanges(pages: PageLayout[]): Map<number, { top: number; bottom: number }> {
  const ranges = new Map<number, { top: number; bottom: number }>();
  for (const page of pages) {
    for (const it of page.items) {
      const r = ranges.get(it.paraIndex);
      if (!r) ranges.set(it.paraIndex, { top: it.y, bottom: it.y + it.h });
      else {
        r.top = Math.min(r.top, it.y);
        r.bottom = Math.max(r.bottom, it.y + it.h);
      }
    }
  }
  return ranges;
}

/** 縦に積んだページ列の中での、段落ごとの絶対位置 (px) */
function stackPositions(pages: PageLayout[], top: number, g: PageGeom): Map<number, { top: number; bottom: number }> {
  const pos = new Map<number, { top: number; bottom: number }>();
  pages.forEach((page, j) => {
    const pageTop = top + j * (g.height + PAGE_GAP) + g.textY;
    for (const it of page.items) {
      const t = pageTop + it.y * g.lineHeight;
      const b = pageTop + (it.y + it.h) * g.lineHeight;
      const r = pos.get(it.paraIndex);
      if (!r) pos.set(it.paraIndex, { top: t, bottom: b });
      else {
        r.top = Math.min(r.top, t);
        r.bottom = Math.max(r.bottom, b);
      }
    }
  });
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

function unitColor(u: SlopeUnit): string {
  switch (u.change) {
    case "deleted":
      return RED;
    case "added":
      return u.addedByBulk ? ORANGE : GREEN;
    case "bulk":
      return ORANGE;
    case "fine":
      return GREEN;
    case "movedFrom":
    case "movedTo":
    case "moved":
      return BLUE;
    default:
      return GRAY;
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
  g: PageGeom
): string {
  const left = stackPositions(s.startPages, stackTop, g);
  const right = stackPositions(s.endPages, stackTop, g);
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
    const color = unitColor(u);
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
    parts.push(`<path d="${d}" fill="rgb(${color})" fill-opacity="${opacity.toFixed(2)}" stroke="rgb(${color})" stroke-opacity="${Math.min(1, opacity + 0.2).toFixed(2)}" stroke-width="0.5"/>`);
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
      `<path d="${band(whole ? L : narrow(L), whole ? R : narrow(R))}" fill="rgb(${BLUE})" fill-opacity="0.55" stroke="rgb(${BLUE})" stroke-opacity="0.9" stroke-width="0.6"/>`
    );
  }
  return parts.join("\n");
}

/** 区間 index (0始まり) のキャプション。x〜x+width は前後の列の中心の間 */
function renderCaption(s: FlowSession, index: number, x: number, y: number, width: number, font: string): string {
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
  if (s.bulkInsChars > 0) notes.push([ORANGE, t("flowBulkNote", fmtCount(s.bulkInsChars))]);
  if (s.movedChars > 0) notes.push([BLUE, t("flowMovedNote", fmtCount(s.movedChars))]);
  notes.forEach(([rgb, text], i) => {
    lines.push(
      `<text x="${cx}" y="${y + 49 + i * 15}" text-anchor="middle" font-size="11" fill="rgb(${rgb})">${esc(text)}</text>`
    );
  });
  return `<g font-family="${font}">${lines.join("")}</g>`;
}

/** 凡例。返り値の width は凡例全体の幅 (図の幅を決めるのに使う) */
function renderLegend(x: number, y: number, font: string): { svg: string; width: number } {
  const rows: [string, string][][] = [
    [
      [GREEN, t("flowLegendFine")],
      [ORANGE, t("flowLegendBulk")],
    ],
    [
      [RED, t("flowLegendDeleted")],
      [BLUE, t("flowLegendMoved")],
      [GRAY, t("flowLegendUnchanged")],
    ],
  ];
  const parts: string[] = [];
  let width = 0;
  rows.forEach((row, r) => {
    let dx = 0;
    const dy = r * 20;
    for (const [rgb, label] of row) {
      parts.push(
        `<rect x="${x + dx}" y="${y + dy - 10}" width="12" height="12" fill="${fill(rgb, 0.3)}"/>` +
          `<rect x="${x + dx + 12}" y="${y + dy - 10}" width="12" height="12" fill="${fill(rgb, 1)}"/>` +
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
  const g0 = result.geometry;
  const pageWidth = opts.pageWidth ?? 150;
  const slopeWidth = opts.slopeWidth ?? 72;
  const scale = pageWidth / g0.pageWidthTwips;
  const geom: PageGeom = {
    width: pageWidth,
    height: g0.pageHeightTwips * scale,
    textX: g0.marginLeftTwips * scale,
    textY: g0.marginTopTwips * scale,
    textWidth: (g0.pageWidthTwips - g0.marginLeftTwips - g0.marginRightTwips) * scale,
    // 1ページの行数 (補正後) でページの本文領域を割り、行がページからはみ出さないようにする
    lineHeight:
      ((g0.pageHeightTwips - g0.marginTopTwips - g0.marginBottomTwips) * scale) / result.linesPerPage,
  };

  // 列: [区間1の開始時点, 区間1の終了時点, 区間2の終了時点, ...]
  const columns: PageLayout[][] = [sessions[0].startPages, ...sessions.map((s) => s.endPages)];
  const columnX = (k: number) => MARGIN + k * (pageWidth + slopeWidth);
  const maxPages = Math.max(...columns.map((c) => c.length));
  const stackTop = HEADER_HEIGHT + STACK_LABEL_HEIGHT;
  const columnsHeight = maxPages * geom.height + (maxPages - 1) * PAGE_GAP;
  const captionY = stackTop + columnsHeight + 20;
  const legend = renderLegend(MARGIN, 58, font);
  const width = Math.ceil(
    Math.max(720, MARGIN * 2 + legend.width, MARGIN * 2 + columns.length * pageWidth + sessions.length * slopeWidth)
  );
  const height = captionY + CAPTION_HEIGHT + MARGIN;

  const parts: string[] = [];
  parts.push(`<rect width="100%" height="100%" fill="#ffffff"/>`);
  parts.push(
    `<text x="${MARGIN}" y="32" font-family="${font}" font-size="18" font-weight="bold" fill="#222">${esc(
      opts.title ?? t("flowDefaultTitle")
    )}</text>`
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
      for (const i of u.paraIndexes) nextMark.set(i, u.change === "deleted" ? RED : BLUE);
    }
    const tintOf = (i: number): Tint | undefined => {
      const t: Tint = { nextMark: nextMark.get(i) };
      const h: ParaHeat | undefined = ended?.heat.get(i);
      if (h) {
        // 最も度合いの大きい色で段落を塗り、2番目の色を左余白の細い帯で示す
        const ranked = [
          { c: GREEN, v: h.green },
          { c: ORANGE, v: h.orange },
          { c: BLUE, v: h.moved },
        ]
          .filter((e) => e.v >= MIN_INTENSITY)
          .sort((a, b) => b.v - a.v);
        if (ranked[0]) Object.assign(t, { main: ranked[0].c, mainV: ranked[0].v });
        if (ranked[1]) Object.assign(t, { sub: ranked[1].c, subV: ranked[1].v });
      }
      return t.main || t.nextMark ? t : undefined;
    };
    pages.forEach((page, j) => {
      parts.push(renderPage(page, tintOf, x, stackTop + j * (geom.height + PAGE_GAP), geom));
    });
  });

  sessions.forEach((s, i) => {
    parts.push(renderSlopes(s, columnX(i) + pageWidth, columnX(i + 1), stackTop, geom));
    parts.push(
      renderCaption(s, i, columnX(i) + pageWidth / 2, captionY, pageWidth + slopeWidth, font)
    );
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height.toFixed(0)}" viewBox="0 0 ${width} ${height.toFixed(0)}">
${parts.join("\n")}${renderNote(opts.note, width, Math.round(height), ` font-family="${font}"`)}
</svg>
`;
}
