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
import { esc } from "./svgChart";
import type { LocalizedText } from "./categories";
import { getLang, Lang } from "./i18n";

export const chartBarKey = (categoryId: string, start: Date) => `chart:bar:${categoryId}:${start.toISOString()}`;
export const flowParaKey = (column: number, paraIndex: number) => `flow:para:${column}:${paraIndex}`;
export const flowBandKey = (session: number, unitKey: string) => `flow:band:${session}:${unitKey}`;
export const flowMoveKey = (session: number, fromKey: string, toKey: string) => `flow:move:${session}:${fromKey}:${toKey}`;
export const flowCaptionKey = (session: number) => `flow:caption:${session}`;

/** 図の部分の情報 (注釈を付ける側が、どの部分に何を付けるかを決めるため) */
export type FigureTarget =
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
  | { key: string; figure: "flow"; kind: "caption"; session: number; start: string; end: string };

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
