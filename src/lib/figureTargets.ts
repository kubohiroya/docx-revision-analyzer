/**
 * figureTargets.ts
 *
 * 図 (chart / flow の SVG) の部分を指すキーと、部分に付ける注釈 (マウスオーバーで出る説明、クリックで開くリンク)。
 *
 * 部分とキー:
 *  - chart の棒:        chart:bar:<カテゴリ id>:<バケットの開始時刻 ISO 8601>  (削除の棒のカテゴリは "deleted")
 *  - flow の段落:        flow:para:<列>:<段落の番号 (Para.index)>  (列 0 は区間1の開始時点、列 k は区間 k の終了時点)
 *  - flow の帯:          flow:band:<区間 (0始まり)>:<単位のキー (p<段落> / t<表>)>
 *  - flow の移動の帯:    flow:move:<区間>:<移動元の単位>:<移動先の単位>
 *  - flow のキャプション: flow:caption:<区間>
 *
 * 注釈は、SVG の要素に <title> (ブラウザの標準のツールチップ) と <a href> (クリックで開くリンク) として埋め込む。
 * 注釈が無く annotate も指定しなければ、SVG は変わらない。
 */

import { Bucket, buildBuckets, BucketSpec } from "./timeBuckets";
import type { Session } from "./sessions";
import type { FlowResult } from "./flow";
import type { DocxLayoutModel, Para, RevRef } from "./docxLayout";
import { esc } from "./svgChart";
import type { LocalizedText } from "./categories";
import { getLang, Lang } from "./i18n";

export const chartBarKey = (categoryId: string, start: Date) => `chart:bar:${categoryId}:${start.toISOString()}`;
export const flowParaKey = (column: number, paraIndex: number) => `flow:para:${column}:${paraIndex}`;
export const flowBandKey = (session: number, unitKey: string) => `flow:band:${session}:${unitKey}`;
export const flowMoveKey = (session: number, fromKey: string, toKey: string) => `flow:move:${session}:${fromKey}:${toKey}`;
export const flowCaptionKey = (session: number) => `flow:caption:${session}`;

/** 部分に付け加える情報 (enrichTargets で設定する) */
export interface TargetExtras {
  /** 段落の冒頭 (その時点の文書の本文、約 40 文字)。段落・帯だけ */
  excerpt?: string;
  /** 段落を含むセクションの名前 (その段落以前で最後の見出しの文字、またはブックマークの名前)。段落・帯だけ */
  section?: string;
  /**
   * 元の文書の該当箇所へのリンク (OneDrive / SharePoint の文書のとき)。見出しのセクションなら Word for the web の
   * 見出しリンク (nav=)、ブックマークなら URL#ブックマーク名、どちらも無ければ文書の URL (冒頭)
   */
  docLink?: string;
}

/** 図の部分の情報 (注釈を付ける側が、どの部分に何を付けるかを決めるため) */
export type FigureTarget = TargetExtras & (
  | { key: string; figure: "chart"; kind: "bar"; categoryId: string; start: string; end: string; chars: number }
  | { key: string; figure: "flow"; kind: "paragraph"; column: number; paraIndex: number; session: number | null }
  | {
      key: string;
      figure: "flow";
      kind: "band";
      session: number;
      unitKey: string;
      paraIndexes: number[];
      change: string;
    }
  | { key: string; figure: "flow"; kind: "move"; session: number; fromKey: string; toKey: string; chars: number }
  | { key: string; figure: "flow"; kind: "caption"; session: number; start: string; end: string }
);

/** 注釈 (言語ごとの文を持つ形)。外部の注釈ファイルや拡張モジュールが作る */
export interface FigureAnnotation {
  /** 部分のキー */
  target: string;
  /** マウスオーバーで出る説明 */
  tooltip?: LocalizedText;
  /** クリックで開くリンク (http / https のみ) */
  href?: string;
}

/** SVG に埋め込む形の注釈 (表示言語の文に解決したもの) */
export interface ResolvedAnnotation {
  title?: string;
  href?: string;
}

function localized(v: LocalizedText | undefined, lang: Lang): string | undefined {
  if (v === undefined) return undefined;
  if (typeof v === "string") return v;
  return v[lang] ?? v.en ?? v.ja;
}

/** リンクとして使える URL か (http / https のみ。javascript: などは使わない) */
export function isSafeHref(href: string): boolean {
  try {
    const u = new URL(href);
    return u.protocol === "https:" || u.protocol === "http:";
  } catch {
    return false;
  }
}

/**
 * 注釈を部分ごとにまとめる。同じ部分に複数の注釈があれば説明は改行でつなぎ、リンクは最初のものを使う。
 * 安全でない URL は捨てる。
 */
export function resolveAnnotations(annotations: FigureAnnotation[], lang: Lang = getLang()): Map<string, ResolvedAnnotation> {
  const map = new Map<string, ResolvedAnnotation>();
  for (const a of annotations) {
    if (typeof a?.target !== "string") continue;
    const r = map.get(a.target) ?? {};
    const title = localized(a.tooltip, lang)?.trim();
    if (title) r.title = r.title ? `${r.title}\n${title}` : title;
    if (!r.href && typeof a.href === "string" && isSafeHref(a.href)) r.href = a.href;
    map.set(a.target, r);
  }
  return map;
}

export interface DecorateOptions {
  /** data-target (部分のキー) を付ける */
  annotate?: boolean;
  annotations?: Map<string, ResolvedAnnotation>;
}

/**
 * SVG の要素 (文字列。"<rect .../>" や "<path ...>...</path>"、"<g>...</g>") に、部分のキー・説明・リンクを付ける。
 * 何も付けないときは元の文字列をそのまま返す。
 */
export function decorateElement(el: string, key: string, o: DecorateOptions): string {
  const ann = o.annotations?.get(key);
  if (!o.annotate && !ann) return el;
  const tag = el.match(/^<([a-zA-Z]+)/)?.[1];
  if (!tag) return el;
  let out = o.annotate ? el.replace(/^<[a-zA-Z]+/, `<${tag} data-target="${esc(key)}"`) : el;
  if (ann?.title) {
    const title = esc(ann.title);
    const head = out.match(/^<[^>]*>/)![0];
    if (head.endsWith("/>")) {
      // 空要素: 中身として <title> を入れる
      out = `${head.slice(0, -2)}><title>${title}</title></${tag}>${out.slice(head.length)}`;
    } else if (new RegExp(`^<[^>]*><title>`).test(out)) {
      // すでに <title> がある (chart の棒の数値の説明など): 後ろに足す
      out = out.replace("</title>", `\n${title}</title>`);
    } else {
      out = `${head}<title>${title}</title>${out.slice(head.length)}`;
    }
  }
  if (ann?.href) out = `<a href="${esc(ann.href)}" target="_blank" rel="noopener noreferrer">${out}</a>`;
  return out;
}

// ---------------------------------------------------------------------------
// 部分の一覧
// ---------------------------------------------------------------------------

/** chart の部分 (棒) の一覧。highlightIds は描いているハイライトのカテゴリ (priority の低い順) */
export function chartTargets(buckets: Bucket[], highlightIds: string[]): FigureTarget[] {
  const out: FigureTarget[] = [];
  for (const b of buckets) {
    const series: [string, number][] = [
      ["fine", b.addedFine],
      ...highlightIds.map((id): [string, number] => [id, b.addedByCategory[id] ?? 0]),
      ["moved", b.addedMoved],
      ["deleted", b.deleted],
    ];
    for (const [id, chars] of series) {
      if (chars <= 0) continue;
      out.push({
        key: chartBarKey(id, b.start),
        figure: "chart",
        kind: "bar",
        categoryId: id,
        start: b.start.toISOString(),
        end: b.end.toISOString(),
        chars,
      });
    }
  }
  return out;
}

/**
 * 無編集期間で分けた chart (renderSessionedRevisionChart) の部分の一覧。
 * 描画と同じく、区間ごとにバケットに分ける
 */
export function sessionedChartTargets(
  sessions: Session[],
  baselineCharCount: number,
  bucketSpec: BucketSpec,
  highlightIds: string[]
): FigureTarget[] {
  let running = baselineCharCount;
  const out: FigureTarget[] = [];
  for (const s of sessions) {
    out.push(...chartTargets(buildBuckets(s.events, running, bucketSpec), highlightIds));
    for (const ev of s.events) running += ev.type === "ins" ? ev.chars : -ev.chars;
  }
  return out;
}

/** flow の部分 (段落・帯・移動の帯・キャプション) の一覧 */
export function flowTargets(result: FlowResult): FigureTarget[] {
  const out: FigureTarget[] = [];
  const s = result.sessions;
  const columns = s.length ? [s[0].startPages, ...s.map((x) => x.endPages)] : [];
  columns.forEach((pages, k) => {
    const seen = new Set<number>();
    for (const page of pages) {
      for (const it of page.items) {
        if (seen.has(it.paraIndex)) continue;
        seen.add(it.paraIndex);
        out.push({
          key: flowParaKey(k, it.paraIndex),
          figure: "flow",
          kind: "paragraph",
          column: k,
          paraIndex: it.paraIndex,
          session: k === 0 ? null : k - 1,
        });
      }
    }
  });
  s.forEach((session, i) => {
    for (const u of session.units) {
      out.push({
        key: flowBandKey(i, u.key),
        figure: "flow",
        kind: "band",
        session: i,
        unitKey: u.key,
        paraIndexes: u.paraIndexes,
        change: u.change,
      });
    }
    for (const l of session.links) {
      out.push({ key: flowMoveKey(i, l.fromKey, l.toKey), figure: "flow", kind: "move", session: i, fromKey: l.fromKey, toKey: l.toKey, chars: l.chars });
    }
    out.push({
      key: flowCaptionKey(i),
      figure: "flow",
      kind: "caption",
      session: i,
      start: session.start.toISOString(),
      end: session.end.toISOString(),
    });
  });
  return out;
}

// ---------------------------------------------------------------------------
// 段落の冒頭・セクション・元の文書へのリンク
// ---------------------------------------------------------------------------

/** 冒頭として示す文字数 */
export const EXCERPT_CHARS = 40;

/** 時刻 t の文書で見えている断片か (flow.ts の表示状態と同じ規則) */
function visibleAt(ins: RevRef | undefined, del: RevRef | undefined, t: number): boolean {
  const inserted = !ins || !ins.date || ins.date.getTime() <= t;
  const deleted = !!del && (!del.date || del.date.getTime() <= t);
  return inserted && !deleted;
}

/** 時刻 t の段落の冒頭 (max 文字を超えたら … で切る) */
export function paragraphExcerpt(p: Para, t: number, max = EXCERPT_CHARS): string {
  let text = "";
  for (const s of p.segs) {
    if (s.kind !== "text" || !visibleAt(s.ins, s.del, t)) continue;
    text += s.text;
    if (text.length > max * 2) break;
  }
  const chars = [...text.replace(/\s+/g, " ").trim()];
  return chars.length > max ? `${chars.slice(0, max).join("")}…` : chars.join("");
}

/**
 * 段落を含むセクション。その段落以前で最後に現れた見出し (見出しの段落そのものを含む) または
 * ブックマークのうち、後に現れた方
 */
export type Section =
  | { kind: "heading"; text: string; paraId?: string; bookmark?: string }
  | { kind: "bookmark"; name: string };

/** セクションの表示名 (見出しの文字またはブックマークの名前) */
export function sectionLabel(s: Section): string {
  return s.kind === "heading" ? s.text : s.name;
}

/** 段落ごとの、その段落を含むセクション */
export function sectionsOf(model: DocxLayoutModel): Map<number, Section> {
  const out = new Map<number, Section>();
  let current: Section | undefined;
  for (const p of model.paras) {
    const bookmark = p.bookmarks?.length ? p.bookmarks[p.bookmarks.length - 1] : undefined;
    if (p.kind === "heading") {
      current = { kind: "heading", text: paragraphExcerpt(p, Infinity), paraId: p.paraId, bookmark };
    } else if (bookmark) {
      current = { kind: "bookmark", name: bookmark };
    }
    if (current) out.set(p.index, current);
  }
  return out;
}

/**
 * Word for the web の見出しリンク: 文書の URL に nav={"h":"<見出しの段落の w14:paraId を10進にしたもの>"} (base64) を付ける。
 * Word for the web の「見出しリンクをコピー」が作る URL と同じ形
 */
export function headingNavUrl(webUrl: string, paraId: string): string {
  const base = webUrl.replace(/#.*$/, "").replace(/([?&])nav=[^&]*&?/, "$1").replace(/[?&]$/, "");
  const nav = btoa(JSON.stringify({ h: String(parseInt(paraId, 16)) }));
  return `${base}${base.includes("?") ? "&" : "?"}nav=${encodeURIComponent(nav)}`;
}

/**
 * 元の文書へのリンク:
 *  - 見出しのセクション (見出しに w14:paraId がある): Word for the web の見出しリンク (nav=)
 *  - ブックマークのセクション: URL#ブックマーク名 (Word デスクトップ向け。Word for the web は冒頭を開く)
 *  - それ以外: 文書の URL (冒頭)
 */
export function documentLink(webUrl: string, section?: Section): string {
  if (section?.kind === "heading" && section.paraId) return headingNavUrl(webUrl, section.paraId);
  const bookmark = section?.kind === "bookmark" ? section.name : section?.bookmark;
  return bookmark ? `${webUrl.replace(/#.*$/, "")}#${encodeURIComponent(bookmark)}` : webUrl.replace(/#.*$/, "");
}

/**
 * flow の部分に、段落の冒頭・セクション・元の文書へのリンク (webUrl があるとき) を付け加える。
 * chart の部分とキャプションには、元の文書へのリンク (冒頭) だけを付ける。
 * 段落の冒頭は、その列 (区間の開始時点 / 終了時点) の文書の本文
 */
export function enrichTargets(
  targets: FigureTarget[],
  model: DocxLayoutModel | undefined,
  flow: FlowResult | undefined,
  opts: { webUrl?: string; excerptChars?: number } = {}
): FigureTarget[] {
  const sections = model ? sectionsOf(model) : new Map<number, Section>();
  const byIndex = new Map((model?.paras ?? []).map((p) => [p.index, p]));
  const sessions = flow?.sessions ?? [];
  /** 列 k の時点 (列 0 は区間1の開始の直前、列 k は区間 k の終了時点) */
  const columnTime = (k: number) => (k === 0 ? (sessions[0]?.start.getTime() ?? 0) - 1 : (sessions[k - 1]?.end.getTime() ?? Infinity));
  const max = opts.excerptChars ?? EXCERPT_CHARS;
  return targets.map((t) => {
    let paraIndex: number | undefined;
    let time = Infinity;
    if (t.kind === "paragraph") {
      paraIndex = t.paraIndex;
      time = columnTime(t.column);
    } else if (t.kind === "band") {
      paraIndex = t.paraIndexes[0];
      // 区間の終了時点で消えている段落 (削除) は、開始時点の本文を示す
      const end = columnTime(t.session + 1);
      const p = paraIndex !== undefined ? byIndex.get(paraIndex) : undefined;
      time = p && paragraphExcerpt(p, end, 1) === "" ? columnTime(t.session) : end;
    }
    const p = paraIndex !== undefined ? byIndex.get(paraIndex) : undefined;
    const section = paraIndex !== undefined ? sections.get(paraIndex) : undefined;
    const extras: TargetExtras = {};
    if (p) {
      const ex = paragraphExcerpt(p, time, max);
      if (ex) extras.excerpt = ex;
      if (section && sectionLabel(section)) extras.section = sectionLabel(section);
    }
    if (opts.webUrl) extras.docLink = documentLink(opts.webUrl, section);
    return { ...t, ...extras };
  });
}
