/**
 * revisionPositions.ts
 *
 * レイアウトモデル (docxLayout.ts の DocxLayoutModel) から、挿入・削除の1件ごとに
 * 「文書のどこで起きたか」を持つ編集イベント列を取り出す。
 *
 * 位置の表し方:
 *  - 文書の状態 (最終文書、または編集の時点の文書) を、存在する段落を "\n" で連結した
 *    1つの文字列とみなし、その中の位置で表す (stateText / finalDocumentText で得られる)。
 *  - 文字数・オフセットは JavaScript の文字列長 (UTF-16 コード単位) で数える。
 *    タブは "\t"、改行・改ページは "\n" として1文字、図は0文字とする
 *    (RevisionEvent.chars の数え方と同じ)。
 *  - 段落インデックスは「その状態で存在する段落」の通し番号 (0始まり)。
 *    モデル全体での番号 (Para.index) は paraModelIndex に入れる。
 *
 * 対応づけ:
 *  - イベントは (type, w:id) ごとに1件。RevisionEvent とは revisionKey で対応づく。
 *    attachRevisionPositions で RevisionEvent.position に位置を設定できる。
 *
 * 移動 (w:moveFrom / w:moveTo):
 *  - moveTo は type "ins"、moveFrom は type "del" のイベントとして move = true で含める。
 *    移動元と移動先は moveName (Word が付ける範囲名) で対応づく。
 *  - RevisionEvent (docxRevisions.ts) は移動を含まないため、移動イベントには対応する
 *    RevisionEvent が無い。
 *
 * 対象外: テキストボックス内・脚注・ヘッダー等 (レイアウトモデルと同じ)。
 */

import type { DocxLayoutModel, Para, RevRef, Seg } from "./docxLayout";
import type { RevisionEvent, RevisionType } from "./docxRevisions";

export interface RevisionPosition {
  /** その状態で存在する段落の中での段落インデックス (0始まり) */
  paraIndex: number;
  /** レイアウトモデル上の段落番号 (Para.index) */
  paraModelIndex: number;
  /** 段落内の文字オフセット */
  offsetInPara: number;
  /** 文書先頭からの文字オフセット (段落の区切りを "\n" 1文字として数える) */
  docOffset: number;
  /**
   * 変更の属する段落がその状態に存在しない場合 true。
   * このとき位置は、その段落があった場所 (次に残っている段落の先頭、
   * 後ろに段落が無ければ文書末尾) を指す。
   */
  paraMissing: boolean;
}

/** 変更が占める範囲。start と end が同じなら、その状態では文字を持たない (削除済み等) */
export interface RevisionRange {
  start: RevisionPosition;
  end: RevisionPosition;
}

export interface PositionedRevisionEvent {
  type: RevisionType;
  /** w:id */
  id: string;
  /** 日時が無い (個人情報削除等) 場合は undefined */
  date?: Date;
  author?: string;
  /** 移動 (w:moveFrom / w:moveTo) による変更か */
  move: boolean;
  /** 移動の範囲名。同じ名前の moveFrom と moveTo が1つの移動の移動元と移動先 */
  moveName?: string;
  /** 挿入/削除された文字数 (段落記号は数えない) */
  chars: number;
  /** 挿入/削除された本文 */
  text: string;
  /** 段落記号 (段落の挿入/削除) を含むか */
  includesParaMark: boolean;
  /** 変更が及ぶ段落のレイアウトモデル上の番号 (Para.index) */
  paraModelIndices: number[];
  /**
   * 最終文書 (すべての変更を反映した状態) での位置。
   * 削除や、後で削除された挿入では、文字が消えた場所 (start = end) を指す。
   */
  final: RevisionRange;
  /** 最終文書に残っている文字数 (削除では 0) */
  finalChars: number;
  /**
   * 編集の時点の文書での位置。挿入では挿入直後、削除では削除直前の状態
   * (その時刻より前の変更だけを反映した状態) で測る。日時が無い場合は undefined。
   */
  atEdit?: RevisionRange;
}

/** RevisionEvent と PositionedRevisionEvent を対応づけるキー */
export function revisionKey(type: RevisionType, id: string | number | undefined): string {
  return `${type}:${id ?? ""}`;
}

// ---------------------------------------------------------------------------
// 文書の状態
// ---------------------------------------------------------------------------

/** 断片が状態の中に見えているかの判定 */
type Visible = (ins: RevRef | undefined, del: RevRef | undefined) => boolean;

/** すべての変更を反映した状態 */
const finalState: Visible = (_ins, del) => !del;

/**
 * 時刻 t の状態 (t 以前の変更を反映)。docx-revision-flow と同じく、日時の無い挿入は
 * 「最初から存在」、日時の無い削除は「最初から削除済み」とみなす。
 */
function stateAt(t: number): Visible {
  return (ins, del) => {
    const inserted = !ins || !ins.date || ins.date.getTime() <= t;
    const deleted = !!del && (!del.date || del.date.getTime() <= t);
    return inserted && !deleted;
  };
}

function segText(s: Seg): string {
  if (s.kind === "text") return s.text;
  if (s.kind === "lineBreak" || s.kind === "pageBreak") return "\n";
  return "";
}

function paraExists(p: Para, visible: Visible): boolean {
  return p.segs.some((s) => visible(s.ins, s.del)) || visible(p.markIns, p.markDel);
}

/** 状態 visible での文書の本文 (存在する段落を "\n" で連結したもの) */
function stateText(model: DocxLayoutModel, visible: Visible): string {
  return model.paras
    .filter((p) => paraExists(p, visible))
    .map((p) => p.segs.filter((s) => visible(s.ins, s.del)).map(segText).join(""))
    .join("\n");
}

/** 最終文書 (すべての変更を反映した状態) の本文。位置のオフセットはこの文字列の中の位置 */
export function finalDocumentText(model: DocxLayoutModel): string {
  return stateText(model, finalState);
}

/** 時刻 date の状態 (date 以前の変更を反映) の本文 */
export function documentTextAt(model: DocxLayoutModel, date: Date): string {
  return stateText(model, stateAt(date.getTime()));
}

// ---------------------------------------------------------------------------
// 位置の計算
// ---------------------------------------------------------------------------

/** 断片 (または段落記号) が対象の変更に属するか */
type Belongs = (ins: RevRef | undefined, del: RevRef | undefined) => boolean;

/** 状態 visible の中で、対象の変更に属する断片が占める範囲を求める */
function locate(model: DocxLayoutModel, visible: Visible, belongs: Belongs): RevisionRange | undefined {
  let start: RevisionPosition | undefined;
  let end: RevisionPosition | undefined;
  let paraIndex = 0;
  let docOffset = 0;
  /** 存在しない段落の変更は、次に存在する段落の先頭で位置を確定する */
  let pendingMissing: Para | undefined;
  /** 直前に存在した段落の長さ (文書末尾の位置を求めるため) */
  let lastParaLength = 0;

  const pos = (p: Para, offsetInPara: number, paraMissing: boolean): RevisionPosition => ({
    paraIndex,
    paraModelIndex: p.index,
    offsetInPara,
    docOffset: docOffset + offsetInPara,
    paraMissing,
  });

  for (const p of model.paras) {
    const hit = p.segs.some((s) => belongs(s.ins, s.del)) || belongs(p.markIns, p.markDel);
    if (!paraExists(p, visible)) {
      if (hit) pendingMissing ??= p;
      continue;
    }
    if (pendingMissing) {
      const at = { ...pos(p, 0, true), paraModelIndex: pendingMissing.index };
      start ??= at;
      end = at;
      pendingMissing = undefined;
    }
    let offset = 0;
    for (const s of p.segs) {
      const shown = visible(s.ins, s.del);
      if (belongs(s.ins, s.del)) {
        start ??= pos(p, offset, false);
        end = pos(p, offset + (shown ? segText(s).length : 0), false);
      }
      if (shown) offset += segText(s).length;
    }
    if (belongs(p.markIns, p.markDel)) {
      start ??= pos(p, offset, false);
      end = pos(p, offset, false);
    }
    docOffset += offset + 1; // 段落の区切り "\n"
    lastParaLength = offset;
    paraIndex++;
  }
  if (pendingMissing) {
    // 後ろに存在する段落が無い: 文書末尾 (最後の段落の終わり)
    const at: RevisionPosition = {
      paraIndex: Math.max(0, paraIndex - 1),
      paraModelIndex: pendingMissing.index,
      offsetInPara: lastParaLength,
      docOffset: Math.max(0, docOffset - 1),
      paraMissing: true,
    };
    start ??= at;
    end = at;
  }
  return start && end ? { start, end } : undefined;
}

// ---------------------------------------------------------------------------
// イベントの抽出
// ---------------------------------------------------------------------------

interface Group {
  type: RevisionType;
  rev: RevRef;
  chars: number;
  text: string;
  finalChars: number;
  includesParaMark: boolean;
  paras: Set<number>;
}

/**
 * レイアウトモデルから、挿入・削除ごとの位置付き編集イベント列を取り出す。
 * 日時の順に並べ (同じ日時なら文書中の出現順)、日時の無いものは末尾に置く。
 */
export function extractRevisionPositions(model: DocxLayoutModel): PositionedRevisionEvent[] {
  const groups = new Map<string, Group>();
  const add = (type: RevisionType, rev: RevRef | undefined, p: Para, s: Seg | undefined) => {
    if (!rev) return;
    const key = revisionKey(type, rev.id);
    let g = groups.get(key);
    if (!g) {
      g = { type, rev, chars: 0, text: "", finalChars: 0, includesParaMark: false, paras: new Set() };
      groups.set(key, g);
    }
    g.paras.add(p.index);
    if (!s) {
      g.includesParaMark = true;
      return;
    }
    const text = segText(s);
    g.chars += text.length;
    g.text += text;
    if (type === "ins" && !s.del) g.finalChars += text.length;
  };
  for (const p of model.paras) {
    for (const s of p.segs) {
      add("ins", s.ins, p, s);
      add("del", s.del, p, s);
    }
    add("ins", p.markIns, p, undefined);
    add("del", p.markDel, p, undefined);
  }

  const events: PositionedRevisionEvent[] = [];
  for (const g of groups.values()) {
    const { type, rev } = g;
    const belongs: Belongs = (ins, del) => {
      const r = type === "ins" ? ins : del;
      return !!r && r.id === rev.id;
    };
    const final = locate(model, finalState, belongs)!;
    let atEdit: RevisionRange | undefined;
    if (rev.date) {
      // 挿入は挿入直後、削除は削除直前 (1ミリ秒前) の状態で測る
      const t = rev.date.getTime() - (type === "del" ? 1 : 0);
      atEdit = locate(model, stateAt(t), belongs);
    }
    events.push({
      type,
      id: rev.id,
      date: rev.date,
      author: rev.author,
      move: rev.move,
      moveName: rev.moveName,
      chars: g.chars,
      text: g.text,
      includesParaMark: g.includesParaMark,
      paraModelIndices: [...g.paras].sort((a, b) => a - b),
      final,
      finalChars: g.finalChars,
      atEdit,
    });
  }

  const time = (e: PositionedRevisionEvent) => e.date?.getTime() ?? Infinity;
  return events.sort((a, b) => time(a) - time(b));
}

/**
 * RevisionEvent (docxRevisions.ts) に、w:id で対応する位置付きイベントの最終文書での
 * 範囲を position として設定する。対応するものが無いイベントはそのまま。
 * 渡した events を書き換えて返す。
 */
export function attachRevisionPositions(
  events: RevisionEvent[],
  positioned: PositionedRevisionEvent[]
): RevisionEvent[] {
  const byKey = new Map(positioned.map((e) => [revisionKey(e.type, e.id), e]));
  for (const e of events) {
    if (e.part !== "word/document.xml") continue;
    const p = byKey.get(revisionKey(e.type, e.id));
    if (p) e.position = p.final;
  }
  return events;
}
