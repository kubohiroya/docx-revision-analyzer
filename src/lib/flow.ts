/**
 * flow.ts
 *
 * 変更履歴を「連続的に編集が行われた時間区間 (セッション)」に分け、区間ごとに
 *  1. 区間開始時点 (最初の変更の直前) と終了時点の文書の状態を復元して、模式的にページへ割り付け、
 *  2. 区間内の編集を段落ごとに集計して、2種類の度合いを求め、
 *  3. 段落・図表 (表は1つの単位) ごとに、開始時点から終了時点への変化を分類する。
 *     - green  (細かい編集): 小さな挿入・削除が繰り返された度合い
 *     - orange (一括挿入)  : 同じ時刻にまとめて挿入された大量の文字が段落に占める割合
 *
 * 一括挿入の判定:
 *   Word は貼り付けた文章を段落ごとに別々の w:ins として記録するが、日時は同じになる。
 *   そこで「同じ作成者・同じ日時」の挿入をひとまとまりとみなし、その合計文字数が
 *   bulkChars 以上なら、含まれる挿入をすべて一括挿入とする。
 *   (Word の日時は分単位で記録されることが多いため、1分間に bulkChars 文字以上
 *    手で入力した場合も一括挿入と判定され得る。既定値はこれを考慮した値にしている)
 *
 * 移動・並べ替えの判定:
 *   - Word が移動として記録したもの (w:moveFrom / w:moveTo) は、範囲名で移動元と移動先を対応づける。
 *   - コピー＋貼り付け＋削除や、移動として記録されなかったカット＋貼り付けは、挿入と削除として
 *     記録される。そこで、挿入された文章 (RELOCATION_MIN_CHARS 文字以上) と同じ内容が文書のどこかで
 *     削除されていれば、文書内での並べ替え・複製とみなし、一括挿入にも細かい編集にも数えない。
 *     同じ区間内で対応する削除が見つかれば、移動元と移動先として対応づける。
 */

import { RevisionEvent } from "./docxRevisions";
import { DocxLayoutModel, Para, RevRef, Seg, PageGeometry } from "./docxLayout";
import { splitIntoSessions } from "./sessions";
import { DEFAULT_BULK_CHARS, matchesCorpus, normalizeForMatch as normalize } from "./insertionKinds";

export interface FlowOptions {
  /** 無編集期間がこれ (時間) を超えたら区間を分ける */
  gapThresholdHours: number;
  /** 同時刻の挿入の合計がこれ以上なら一括挿入とみなす (文字数) */
  bulkChars: number;
  /** 対象期間 (この範囲の変更だけを区間分け・集計に使う) */
  from?: Date;
  to?: Date;
}

export const DEFAULT_FLOW_OPTIONS: FlowOptions = {
  gapThresholdHours: 1,
  bulkChars: DEFAULT_BULK_CHARS,
};

export interface ParaHeat {
  /** 細かい編集の度合い (0〜1) */
  green: number;
  /** 一括挿入の度合い (0〜1) */
  orange: number;
  /** 移動・並べ替えで入ってきた文章が段落に占める割合 (0〜1) */
  moved: number;
  insChars: number;
  delChars: number;
  bulkInsChars: number;
  /** 移動・並べ替えで入ってきた / 出ていった文字数 */
  movedInChars: number;
  movedOutChars: number;
}

/** ページ内の描画要素。y / h は「本文の行」単位 */
export type PageItem =
  | { type: "line"; paraIndex: number; y: number; h: number; widthFrac: number; style: Para["kind"] }
  | { type: "figure"; paraIndex: number; y: number; h: number };

export interface PageLayout {
  items: PageItem[];
}

/**
 * 開始時点から終了時点への、段落・図表単位の変化。
 *  - unchanged: 両方にあり、区間内の編集なし
 *  - fine / bulk: 両方にあり、区間内に編集あり (細かい編集 / 一括挿入が優勢)
 *  - added: 終了時点にだけある (区間内に追加された)
 *  - deleted: 開始時点にだけある (区間内に削除された)
 *  - movedTo / movedFrom: 終了時点 / 開始時点にだけあり、内容の大半が移動・並べ替えによるもの
 *  - moved: 両方にあり、区間内の変化が移動・並べ替えだけ
 */
export type UnitChange =
  | "unchanged"
  | "fine"
  | "bulk"
  | "added"
  | "deleted"
  | "movedTo"
  | "movedFrom"
  | "moved";

/** 区間内の移動・並べ替えの、移動元の単位と移動先の単位の対応 */
export interface SlopeLink {
  fromKey: string;
  toKey: string;
  /** move: Word が移動として記録したもの / relocate: 同じ内容の削除と挿入から推定したもの */
  kind: "move" | "relocate";
  chars: number;
}

export interface SlopeUnit {
  /** 段落なら "p<index>"、表なら "t<tableId>" */
  key: string;
  paraIndexes: number[];
  change: UnitChange;
  /** 追加された単位の内容が一括挿入によるものか */
  addedByBulk: boolean;
  /** 色の濃さ (0〜1) */
  intensity: number;
}

export interface FlowSession {
  start: Date;
  end: Date;
  /** 直前の区間との間の無編集期間 (時間)。先頭は null */
  gapBeforeHours: number | null;
  insChars: number;
  delChars: number;
  bulkInsChars: number;
  /** 区間開始時点 (最初の変更の直前) の文書 */
  startPages: PageLayout[];
  /** 区間終了時点の文書 */
  endPages: PageLayout[];
  /** 移動・並べ替えで移された文字数 */
  movedChars: number;
  heat: Map<number, ParaHeat>;
  units: SlopeUnit[];
  links: SlopeLink[];
}

export interface FlowResult {
  geometry: PageGeometry;
  linesPerPage: number;
  sessions: FlowSession[];
}

// ---------------------------------------------------------------------------
// 時刻 T における表示状態
// ---------------------------------------------------------------------------

/** 日時の無い挿入は「最初から存在」、日時の無い削除は「最初から削除済み」とみなす */
function visibleAt(ins: RevRef | undefined, del: RevRef | undefined, t: number): boolean {
  const inserted = !ins || !ins.date || ins.date.getTime() <= t;
  const deleted = !!del && (!del.date || del.date.getTime() <= t);
  return inserted && !deleted;
}

function paraExistsAt(p: Para, visibleSegs: Seg[], t: number): boolean {
  return visibleSegs.length > 0 || visibleAt(p.markIns, p.markDel, t);
}

// ---------------------------------------------------------------------------
// 模式的なページ割り付け
// ---------------------------------------------------------------------------

/** 見出しの文字の大きさ (本文比) */
function headingScale(level: number | undefined): number {
  if (level === undefined) return 1;
  return level === 0 ? 1.6 : level === 1 ? 1.35 : 1.15;
}

type Block =
  | { type: "text"; lines: number[] } // 各行の幅 (0〜1)
  | { type: "figure"; h: number }
  | { type: "pageBreak" };

/** 表示されている断片を、行・図・改ページの列に変換する */
function toBlocks(segs: Seg[], charsPerLine: number, lineHeightTwips: number): Block[] {
  const blocks: Block[] = [];
  let width = 0;
  let hasText = false;
  const flushText = (force: boolean) => {
    if (!hasText && !force) return;
    const lines: number[] = [];
    let rest = width;
    do {
      lines.push(Math.min(1, rest / charsPerLine));
      rest -= charsPerLine;
    } while (rest > 0);
    blocks.push({ type: "text", lines });
    width = 0;
    hasText = false;
  };
  for (const s of segs) {
    if (s.kind === "text") {
      width += s.width;
      hasText = true;
    } else if (s.kind === "lineBreak") {
      flushText(true);
    } else if (s.kind === "pageBreak") {
      flushText(false);
      blocks.push({ type: "pageBreak" });
    } else if (s.kind === "figure") {
      flushText(false);
      blocks.push({ type: "figure", h: s.heightTwips / lineHeightTwips });
    }
  }
  // 空段落も1行分の高さを占める
  if (hasText || blocks.length === 0) flushText(true);
  return blocks;
}

class PageCursor {
  pages: PageLayout[] = [{ items: [] }];
  y = 0;
  constructor(private readonly linesPerPage: number) {}
  get page(): PageLayout {
    return this.pages[this.pages.length - 1];
  }
  newPage(): void {
    this.pages.push({ items: [] });
    this.y = 0;
  }
  /** 高さ h の要素を置く位置を確保する (入りきらなければ改ページ) */
  reserve(h: number): number {
    if (this.y > 0 && this.y + h > this.linesPerPage) this.newPage();
    const y = this.y;
    this.y += h;
    return y;
  }
}

function layoutAt(
  model: DocxLayoutModel,
  t: number,
  charsPerLine: number,
  linesPerPage: number
): { pages: PageLayout[]; visibleChars: Map<number, number> } {
  const cursor = new PageCursor(linesPerPage);
  const visibleChars = new Map<number, number>();
  const lh = model.geometry.lineHeightTwips;

  for (const p of model.paras) {
    const segs = p.segs.filter((s) => visibleAt(s.ins, s.del, t));
    if (!paraExistsAt(p, segs, t)) continue;
    visibleChars.set(
      p.index,
      segs.reduce((n, s) => n + (s.kind === "text" ? s.chars : 0), 0)
    );

    if (p.pageBreakBefore && cursor.y > 0) cursor.newPage();
    const scale = headingScale(p.headingLevel);
    if (p.kind === "heading" && cursor.y > 0) cursor.y += 0.5; // 見出し前の空き

    for (const b of toBlocks(segs, charsPerLine / scale, lh)) {
      if (b.type === "pageBreak") {
        cursor.newPage();
      } else if (b.type === "figure") {
        const h = Math.min(linesPerPage, Math.max(1, b.h));
        const y = cursor.reserve(h);
        cursor.page.items.push({ type: "figure", paraIndex: p.index, y, h });
      } else {
        for (const w of b.lines) {
          const y = cursor.reserve(scale);
          cursor.page.items.push({
            type: "line",
            paraIndex: p.index,
            y,
            h: scale,
            widthFrac: w,
            style: p.kind,
          });
        }
      }
    }
    if (p.sectionBreakAfter) cursor.newPage();
  }
  // 末尾の空ページ (最後の段落の直後の改ページ等) は除く
  const pages = cursor.pages.filter((pg, i) => i === 0 || pg.items.length > 0);
  return { pages, visibleChars };
}

// ---------------------------------------------------------------------------
// 区間分けと編集の集計
// ---------------------------------------------------------------------------

interface DatedRev {
  type: "ins" | "del";
  rev: RevRef;
  chars: number;
}

/** 日時付きの変更を (type, id) ごとにまとめる (区間分けに使うため、移動も含める) */
function collectRevisions(model: DocxLayoutModel): DatedRev[] {
  const map = new Map<string, DatedRev>();
  const add = (type: "ins" | "del", rev: RevRef | undefined, chars: number) => {
    if (!rev || !rev.date) return;
    const key = `${type}:${rev.id}:${rev.date.getTime()}`;
    const r = map.get(key) ?? { type, rev, chars: 0 };
    r.chars += chars;
    map.set(key, r);
  };
  for (const p of model.paras) {
    for (const s of p.segs) {
      const chars = s.kind === "text" ? s.chars : 0;
      add("ins", s.ins, chars);
      add("del", s.del, chars);
    }
    add("ins", p.markIns, 0);
    add("del", p.markDel, 0);
  }
  return [...map.values()];
}

/** 区間内の変更か (move = true なら移動のみ、false なら移動以外のみ) */
function inRange(rev: RevRef | undefined, start: number, end: number, move = false): rev is RevRef {
  if (!rev || rev.move !== move || !rev.date) return false;
  const t = rev.date.getTime();
  return t >= start && t <= end;
}

// ---------------------------------------------------------------------------
// 並べ替え (同じ内容の削除と挿入) の判定
// ---------------------------------------------------------------------------


interface Piece {
  para: number;
  rev: RevRef;
  text: string;
}

/** 段落ごと・変更ごと (w:id) に、挿入または削除された文章をまとめる */
function collectPieces(model: DocxLayoutModel, type: "ins" | "del"): Piece[] {
  const pieces = new Map<string, Piece>();
  for (const p of model.paras) {
    for (const s of p.segs) {
      const rev = s[type];
      if (s.kind !== "text" || !rev || rev.move) continue;
      const key = `${p.index}|${rev.id}|${rev.date?.getTime()}`;
      const piece = pieces.get(key) ?? { para: p.index, rev, text: "" };
      piece.text += s.text;
      pieces.set(key, piece);
    }
  }
  return [...pieces.values()];
}

interface Relocations {
  /** 並べ替えと判定した挿入・削除 ("段落|id|日時") */
  ins: Set<string>;
  del: Set<string>;
  insPieces: Piece[];
  delPieces: Piece[];
}

const pieceKey = (para: number, rev: RevRef) => `${para}|${rev.id}|${rev.date?.getTime()}`;

/**
 * 文書全体 (全期間) で、挿入された文章と同じ内容が削除されている (またはその逆の) ものを並べ替えとみなす。
 * 削除と挿入が別の区間でも判定する (例: ある区間で複製し、後の区間で元を削除した場合)。
 */
function findRelocations(model: DocxLayoutModel): Relocations {
  // 挿入してから削除された文章 (ins と del の両方を持つ断片) は、自分自身と一致してしまうため除く
  const corpus = (type: "ins" | "del") =>
    model.paras
      .flatMap((p) => p.segs)
      .filter((s) => s.kind === "text" && s[type] && !s[type]!.move && !s[type === "ins" ? "del" : "ins"])
      .map((s) => normalize((s as { text: string }).text))
      .join("");
  const insCorpus = corpus("ins");
  const delCorpus = corpus("del");
  const insPieces = collectPieces(model, "ins");
  const delPieces = collectPieces(model, "del");
  const pick = (pieces: Piece[], other: string) =>
    new Set(
      pieces
        .filter((pc) => matchesCorpus(pc.text, other))
        .map((pc) => pieceKey(pc.para, pc.rev))
    );
  return { ins: pick(insPieces, delCorpus), del: pick(delPieces, insCorpus), insPieces, delPieces };
}

// ---------------------------------------------------------------------------
// 区間ごとの集計
// ---------------------------------------------------------------------------

interface SessionHeat {
  heat: Map<number, ParaHeat>;
  insChars: number;
  delChars: number;
  bulkInsChars: number;
  movedChars: number;
  /** 段落単位の移動元 → 移動先 */
  paraLinks: { from: number[]; to: number[]; kind: SlopeLink["kind"]; chars: number }[];
}

function computeHeat(
  model: DocxLayoutModel,
  start: number,
  end: number,
  visibleChars: Map<number, number>,
  bulkChars: number,
  reloc: Relocations
): SessionHeat {
  const isRelocIns = (p: Para, r: RevRef) => reloc.ins.has(pieceKey(p.index, r));
  const isRelocDel = (p: Para, r: RevRef) => reloc.del.has(pieceKey(p.index, r));

  // 同じ作成者・同じ日時の挿入の合計文字数 (並べ替えは除く) / 削除ごとの合計文字数
  const insGroup = new Map<string, number>();
  const delTotal = new Map<string, number>();
  const groupKey = (r: RevRef) => `${r.author ?? ""}|${r.date!.getTime()}`;
  for (const p of model.paras) {
    for (const s of p.segs) {
      if (s.kind !== "text") continue;
      if (inRange(s.ins, start, end) && !isRelocIns(p, s.ins)) {
        insGroup.set(groupKey(s.ins), (insGroup.get(groupKey(s.ins)) ?? 0) + s.chars);
      }
      if (inRange(s.del, start, end)) delTotal.set(s.del.id, (delTotal.get(s.del.id) ?? 0) + s.chars);
    }
  }

  const heat = new Map<number, ParaHeat>();
  let insChars = 0;
  let delChars = 0;
  let bulkInsChars = 0;
  let movedChars = 0;
  // Word の移動: 範囲名ごとの移動元・移動先の段落
  const moves = new Map<string, { from: Set<number>; to: Set<number>; chars: number }>();
  const moveOf = (name: string) => {
    const m = moves.get(name) ?? { from: new Set<number>(), to: new Set<number>(), chars: 0 };
    moves.set(name, m);
    return m;
  };

  for (const p of model.paras) {
    let pIns = 0;
    let pDel = 0;
    let pBulk = 0;
    let movedIn = 0;
    let movedOut = 0;
    let fineChars = 0;
    const fineIds = new Set<string>();
    for (const s of p.segs) {
      if (s.kind !== "text") continue;
      if (inRange(s.ins, start, end)) {
        pIns += s.chars;
        if (isRelocIns(p, s.ins)) {
          movedIn += s.chars;
        } else if ((insGroup.get(groupKey(s.ins)) ?? 0) >= bulkChars) {
          pBulk += s.chars;
        } else {
          fineChars += s.chars;
          fineIds.add(`ins:${s.ins.id}`);
        }
      }
      if (inRange(s.del, start, end)) {
        pDel += s.chars;
        if (isRelocDel(p, s.del)) {
          movedOut += s.chars;
        } else {
          // 大きな削除や、一括挿入と同時刻の削除 (置き換え前の文章の削除) は細かい編集に数えない
          const replacing = (insGroup.get(groupKey(s.del)) ?? 0) >= bulkChars;
          if (!replacing && (delTotal.get(s.del.id) ?? 0) < bulkChars) {
            fineChars += s.chars;
            fineIds.add(`del:${s.del.id}`);
          }
        }
      }
      if (inRange(s.ins, start, end, true)) {
        movedIn += s.chars;
        moveOf(s.ins.moveName ?? s.ins.id).to.add(p.index);
        moveOf(s.ins.moveName ?? s.ins.id).chars += s.chars;
      }
      if (inRange(s.del, start, end, true)) {
        movedOut += s.chars;
        moveOf(s.del.moveName ?? s.del.id).from.add(p.index);
      }
    }
    insChars += pIns;
    delChars += pDel;
    bulkInsChars += pBulk;
    movedChars += movedIn;
    if (pIns === 0 && pDel === 0 && movedIn === 0 && movedOut === 0) continue;

    const visible = visibleChars.get(p.index) ?? 0;
    heat.set(p.index, {
      green: fineChars > 0 ? Math.min(1, Math.max(fineChars / Math.max(visible, 40), fineIds.size / 8)) : 0,
      orange: visible > 0 ? Math.min(1, pBulk / visible) : 0,
      moved: visible > 0 ? Math.min(1, movedIn / visible) : 0,
      insChars: pIns,
      delChars: pDel,
      bulkInsChars: pBulk,
      movedInChars: movedIn,
      movedOutChars: movedOut,
    });
  }

  const paraLinks: SessionHeat["paraLinks"] = [];
  for (const m of moves.values()) {
    if (m.from.size > 0 && m.to.size > 0) {
      paraLinks.push({ from: [...m.from], to: [...m.to], kind: "move", chars: m.chars });
    }
  }
  // 並べ替え: 区間内の挿入と、区間内の同じ内容の削除を対応づける
  const inSession = (pc: Piece) => pc.rev.date !== undefined && inRange(pc.rev, start, end);
  const delsHere = reloc.delPieces.filter((pc) => inSession(pc) && reloc.del.has(pieceKey(pc.para, pc.rev)));
  for (const ins of reloc.insPieces) {
    if (!inSession(ins) || !reloc.ins.has(pieceKey(ins.para, ins.rev))) continue;
    const n = normalize(ins.text);
    const from = delsHere
      .filter((d) => {
        const dn = normalize(d.text);
        return n.includes(dn) || dn.includes(n);
      })
      .map((d) => d.para);
    if (from.length > 0) {
      paraLinks.push({ from: [...new Set(from)], to: [ins.para], kind: "relocate", chars: [...ins.text].length });
    }
  }
  return { heat, insChars, delChars, bulkInsChars, movedChars, paraLinks };
}

const unitKeyOf = (p: Para) => (p.tableId !== undefined ? `t${p.tableId}` : `p${p.index}`);

/** 段落・図表の単位ごとに、開始時点から終了時点への変化を分類する */
function classifyUnits(
  model: DocxLayoutModel,
  atStart: Map<number, number>,
  atEnd: Map<number, number>,
  heat: Map<number, ParaHeat>
): SlopeUnit[] {
  const groups = new Map<string, number[]>();
  for (const p of model.paras) {
    if (!atStart.has(p.index) && !atEnd.has(p.index)) continue;
    const key = unitKeyOf(p);
    const list = groups.get(key) ?? [];
    list.push(p.index);
    groups.set(key, list);
  }
  const units: SlopeUnit[] = [];
  for (const [key, paraIndexes] of groups) {
    const inStart = paraIndexes.some((i) => atStart.has(i));
    const inEnd = paraIndexes.some((i) => atEnd.has(i));
    // 単位全体の編集量 (表は含まれる段落の合計)
    let fine = 0;
    let bulk = 0;
    let movedIn = 0;
    let movedOut = 0;
    let deletedChars = 0;
    let visible = 0;
    for (const i of paraIndexes) {
      visible += atEnd.get(i) ?? 0;
      const h = heat.get(i);
      if (!h) continue;
      bulk += h.bulkInsChars;
      fine += h.insChars - h.bulkInsChars - h.movedInChars + h.delChars - h.movedOutChars;
      movedIn += h.movedInChars;
      movedOut += h.movedOutChars;
      deletedChars += h.delChars;
    }
    fine = Math.max(0, fine);
    const edited = fine > 0 || bulk > 0;
    const byBulk = bulk > 0 && bulk >= fine;
    const maxHeat = Math.max(0, ...paraIndexes.map((i) => {
      const h = heat.get(i);
      return h ? Math.max(h.green, h.orange) : 0;
    }));
    let change: UnitChange;
    if (inStart && !inEnd) {
      // 削除された文字の半分以上が移動・並べ替えで出ていったものなら移動元
      change = movedOut > 0 && movedOut * 2 >= Math.max(deletedChars, movedOut) ? "movedFrom" : "deleted";
    } else if (!inStart && inEnd) {
      change = movedIn > 0 && movedIn * 2 >= visible ? "movedTo" : "added";
    } else if (!edited) {
      change = movedIn > 0 || movedOut > 0 ? "moved" : "unchanged";
    } else {
      change = byBulk ? "bulk" : "fine";
    }
    units.push({
      key,
      paraIndexes,
      change,
      addedByBulk: change === "added" && byBulk,
      intensity:
        change === "bulk"
          ? Math.min(1, bulk / Math.max(visible, 1))
          : change === "fine"
            ? maxHeat
            : 1,
    });
  }
  return units;
}

/** 段落単位の対応を、段落・図表の単位の対応にまとめる */
function toUnitLinks(model: DocxLayoutModel, paraLinks: SessionHeat["paraLinks"]): SlopeLink[] {
  const links = new Map<string, SlopeLink>();
  for (const l of paraLinks) {
    for (const f of l.from) {
      for (const t of l.to) {
        const fromKey = unitKeyOf(model.paras[f]);
        const toKey = unitKeyOf(model.paras[t]);
        const key = `${fromKey}>${toKey}`;
        const link = links.get(key) ?? { fromKey, toKey, kind: l.kind, chars: 0 };
        link.chars += l.chars;
        links.set(key, link);
      }
    }
  }
  return [...links.values()];
}

export function buildFlow(model: DocxLayoutModel, opts: FlowOptions): FlowResult {
  const g = model.geometry;
  const textWidth = g.pageWidthTwips - g.marginLeftTwips - g.marginRightTwips;
  const textHeight = g.pageHeightTwips - g.marginTopTwips - g.marginBottomTwips;
  const charsPerLine = Math.max(10, Math.floor(textWidth / (g.fontSizePt * 20)));
  let linesPerPage = Math.max(5, Math.floor(textHeight / g.lineHeightTwips));

  // Word が保存時に数えたページ数に、最終状態の見積もりページ数が合うよう1ページの行数を補正する
  // (段落間隔・フォント・禁則処理などを模式図では再現しないため)。
  // 補正は 0.5〜2 倍の範囲に限り、明らかに合わない値に引っ張られないようにする。
  if (model.savedPageCount) {
    for (let i = 0; i < 3; i++) {
      const estimated = layoutAt(model, Infinity, charsPerLine, linesPerPage).pages.length;
      if (estimated === model.savedPageCount) break;
      const base = Math.max(5, Math.floor(textHeight / g.lineHeightTwips));
      const next = Math.round(linesPerPage * (estimated / model.savedPageCount));
      linesPerPage = Math.min(base * 2, Math.max(Math.ceil(base / 2), next));
    }
  }

  const from = opts.from?.getTime() ?? -Infinity;
  const to = opts.to?.getTime() ?? Infinity;
  const events: RevisionEvent[] = collectRevisions(model)
    .filter((r) => r.rev.date!.getTime() >= from && r.rev.date!.getTime() <= to)
    .map((r) => ({
      type: r.type,
      id: r.rev.id,
      author: r.rev.author,
      date: r.rev.date!,
      chars: r.chars,
      text: "",
      part: "word/document.xml",
    }));

  const reloc = findRelocations(model);
  const sessions: FlowSession[] = splitIntoSessions(events, opts.gapThresholdHours).map((s) => {
    const start = s.events[0].date.getTime();
    const end = s.events[s.events.length - 1].date.getTime();
    // 開始時点は区間の最初の変更の直前 (同じ時刻の変更は区間に含まれるため 1ms 前)
    const before = layoutAt(model, start - 1, charsPerLine, linesPerPage);
    const after = layoutAt(model, end, charsPerLine, linesPerPage);
    const h = computeHeat(model, start, end, after.visibleChars, opts.bulkChars, reloc);
    return {
      start: new Date(start),
      end: new Date(end),
      gapBeforeHours: s.gapBeforeHours,
      insChars: h.insChars,
      delChars: h.delChars,
      bulkInsChars: h.bulkInsChars,
      movedChars: h.movedChars,
      startPages: before.pages,
      endPages: after.pages,
      heat: h.heat,
      units: classifyUnits(model, before.visibleChars, after.visibleChars, h.heat),
      links: toUnitLinks(model, h.paraLinks),
    };
  });

  return { geometry: g, linesPerPage, sessions };
}
