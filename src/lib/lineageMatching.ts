import type { LineageKind, LineageLink, LineageRange, MatchEvidence, VersionParagraph } from "./lineage";

export interface Unit {
  order: number;
  range: LineageRange;
  text: string;
  key: string;
  joinedKey: string;
  heading: string;
  kind: VersionParagraph["kind"];
  paraId?: string;
  grams: Set<string>;
  words?: Set<string>;
}

function gramsOf(text: string): Set<string> {
  const chars = Array.from(text);
  const result = new Set<string>();
  for (let i = 0; i + 2 < chars.length; i++) result.add(chars[i] + chars[i + 1] + chars[i + 2]);
  return result;
}

export function prepareUnits(
  paragraphs: VersionParagraph[],
  version: number,
  sentences?: Intl.Segmenter,
  words?: Intl.Segmenter,
): Unit[] {
  const result: Unit[] = [];
  let heading = "";
  for (const p of paragraphs) {
    if (p.kind === "heading") heading = p.text.trim().normalize();
    const parts = sentences ? sentences.segment(p.text) : [{ index: 0, segment: p.text }];
    for (const { index, segment } of parts) {
      const text = segment.trim();
      if (!text) continue;
      const start = index + segment.length - segment.trimStart().length;
      const key = text.normalize().replace(/\s+/gu, " ");
      const joinedKey = key.replace(/\s+/gu, "");
      const wordSet = words && /[\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}]/u.test(key)
        ? new Set(Array.from(words.segment(key)).filter((token) => token.isWordLike).map((token) => token.segment))
        : undefined;
      result.push({
        order: result.length, text, key, joinedKey, heading, kind: p.kind, paraId: p.paraId,
        range: { version, paragraph: p.index, start, end: start + text.length },
        grams: gramsOf(joinedKey), words: wordSet,
      });
    }
  }
  return result;
}

function grouped<T>(items: readonly T[], key: (item: T) => string): Map<string, T[]> {
  const result = new Map<string, T[]>();
  for (const item of items) {
    const k = key(item);
    const list = result.get(k);
    if (list) list.push(item);
    else result.set(k, [item]);
  }
  return result;
}

function overlap(a: Set<string>, b: Set<string>): number {
  let shared = 0;
  for (const value of a.size < b.size ? a : b) if ((a.size < b.size ? b : a).has(value)) shared++;
  return a.size + b.size ? (2 * shared) / (a.size + b.size) : 0;
}

interface Pair {
  from: Unit;
  to: Unit;
  score: number;
  evidence: MatchEvidence[];
}

function bestBy(pairs: Pair[], key: (pair: Pair) => Unit, margin: number): Map<Unit, Pair> {
  const ordered = [...pairs].sort((a, b) => b.score - a.score);
  const result = new Map<Unit, Pair>();
  const visited = new Set<Unit>();
  for (const pair of ordered) {
    const k = key(pair);
    if (!visited.has(k)) {
      visited.add(k);
      result.set(k, pair);
    } else if (result.has(k) && result.get(k)!.score - pair.score < margin) result.delete(k);
  }
  return result;
}

function mutual(pairs: Pair[], margin: number): Pair[] {
  const byFrom = bestBy(pairs, (p) => p.from, margin);
  const byTo = bestBy(pairs, (p) => p.to, margin);
  return [...byFrom.values()].filter((p) => byTo.get(p.to) === p);
}

interface Character {
  value: string;
  unit: Unit;
  start: number;
  end: number;
}

function characters(units: Unit[]): Character[] {
  return units.flatMap((unit) => {
    let start = unit.range.start;
    return Array.from(unit.text, (value) => {
      const c = { value, unit, start, end: start + value.length };
      start = c.end;
      return c;
    });
  });
}

function unchangedRanges(from: Unit[], to: Unit[]): Pick<LineageLink, "unchanged" | "alignmentPartial"> {
  if (from.length === 1 && to.length === 1 && from[0].text === to[0].text) {
    return { unchanged: [{ from: from[0].range, to: to[0].range }] };
  }
  const a = characters(from);
  const b = characters(to);
  const unchanged: LineageLink["unchanged"] = [];
  function add(x: Character, y: Character) {
    const last = unchanged.at(-1);
    if (last && last.from.paragraph === x.unit.range.paragraph && last.to.paragraph === y.unit.range.paragraph
      && last.from.end === x.start && last.to.end === y.start) {
      last.from.end = x.end;
      last.to.end = y.end;
    } else {
      unchanged.push({
        from: { ...x.unit.range, start: x.start, end: x.end },
        to: { ...y.unit.range, start: y.start, end: y.end },
      });
    }
  }
  let start = 0;
  let aEnd = a.length;
  let bEnd = b.length;
  while (start < aEnd && start < bEnd && a[start].value === b[start].value) {
    add(a[start], b[start]);
    start++;
  }
  while (aEnd > start && bEnd > start && a[aEnd - 1].value === b[bEnd - 1].value) {
    aEnd--;
    bEnd--;
  }
  const rows = aEnd - start;
  const cols = bEnd - start;
  const alignmentPartial = rows * cols > 250_000;
  if (rows && cols && !alignmentPartial) {
    const stride = cols + 1;
    const lengths = new Uint16Array((rows + 1) * stride);
    for (let i = rows - 1; i >= 0; i--) {
      for (let j = cols - 1; j >= 0; j--) {
        const cell = i * stride + j;
        lengths[cell] = a[start + i].value === b[start + j].value
          ? lengths[cell + stride + 1] + 1
          : Math.max(lengths[cell + stride], lengths[cell + 1]);
      }
    }
    for (let i = 0, j = 0; i < rows && j < cols;) {
      if (a[start + i].value === b[start + j].value) add(a[start + i++], b[start + j++]);
      else if (lengths[(i + 1) * stride + j] >= lengths[i * stride + j + 1]) i++;
      else j++;
    }
  }
  for (let i = aEnd, j = bEnd; i < a.length; i++, j++) add(a[i], b[j]);
  return { unchanged, ...(alignmentPartial ? { alignmentPartial } : {}) };
}

function link(kind: LineageKind, from: Unit[], to: Unit[], evidence: MatchEvidence[] = []): LineageLink {
  return {
    kind, from: from.map((u) => u.range), to: to.map((u) => u.range), evidence,
    ...(from.length && to.length && kind !== "ambiguous" ? unchangedRanges(from, to) : { unchanged: [] }),
  };
}

/** Returns exact contiguous groups of two or three units under the same heading. */
function adjacentGroups(units: Unit[], used: Set<Unit>): Unit[][] {
  const groups: Unit[][] = [];
  for (let i = 0; i < units.length - 1; i++) {
    const group: Unit[] = [];
    for (const unit of units.slice(i, i + 3)) {
      if (unit.kind !== "body" || unit.heading !== units[i].heading || used.has(unit)) break;
      group.push(unit);
      if (group.length >= 2) groups.push([...group]);
    }
  }
  return groups;
}

interface GroupMatch {
  from: Unit[];
  to: Unit[];
  kind: "split" | "merged";
}

function groupMatches(
  before: Unit[], after: Unit[], usedBefore: Set<Unit>, usedAfter: Set<Unit>,
): { matches: GroupMatch[]; conflicts: GroupMatch[] } {
  const candidates: GroupMatch[] = [];
  function find(single: Unit[], many: Unit[], usedSingle: Set<Unit>, usedMany: Set<Unit>, split: boolean) {
    const groups = grouped(adjacentGroups(many, usedMany), (g) => g.map((u) => u.joinedKey).join(""));
    for (const u of single) {
      if (u.kind !== "body" || u.joinedKey.length < 8 || usedSingle.has(u)) continue;
      for (const g of groups.get(u.joinedKey) ?? []) {
        candidates.push(split ? { from: [u], to: g, kind: "split" } : { from: g, to: [u], kind: "merged" });
      }
    }
  }
  find(before, after, usedBefore, usedAfter, true);
  find(after, before, usedAfter, usedBefore, false);
  const counts = new Map<Unit, number>();
  for (const group of candidates) {
    for (const u of [...group.from, ...group.to]) counts.set(u, (counts.get(u) ?? 0) + 1);
  }
  const matches: GroupMatch[] = [];
  const conflicts: GroupMatch[] = [];
  for (const group of candidates) {
    const unique = [...group.from, ...group.to].every((u) => counts.get(u) === 1);
    (unique ? matches : conflicts).push(group);
  }
  return { matches, conflicts };
}

export function compareUnits(before: Unit[], after: Unit[], past: ReadonlyMap<string, Unit[]>): LineageLink[] {
  const links: LineageLink[] = [];
  const usedBefore = new Set<Unit>();
  const usedAfter = new Set<Unit>();
  const pairs = new Map<Unit, { to: Unit; link: LineageLink }>();
  const fromKeys = grouped(before, (u) => u.key);
  const toKeys = grouped(after, (u) => u.key);
  function add(kind: LineageKind, from: Unit[], to: Unit[], evidence: MatchEvidence[] = []): LineageLink {
    from.forEach((u) => usedBefore.add(u));
    to.forEach((u) => usedAfter.add(u));
    const relationship = link(kind, from, to, evidence);
    links.push(relationship);
    return relationship;
  }
  function pair(p: Pair) {
    pairs.set(p.from, { to: p.to, link: add(p.from.text === p.to.text ? "retained" : "rewritten", [p.from], [p.to], p.evidence) });
  }
  for (const [key, left] of fromKeys) {
    const right = toKeys.get(key);
    if (left.length === 1 && right?.length === 1) pair({ from: left[0], to: right[0], score: 1, evidence: ["exact"] });
  }

  function neighbors(a: Unit, b: Unit): number {
    return [-1, 1].filter((offset) => {
      const local = pairs.get(before[a.order + offset]);
      return local && local.to === after[b.order + offset];
    }).length;
  }

  const identifiedText = (u: Unit) => `${u.paraId ?? ""}\u0000${u.key}`;
  const leftIds = grouped(before.filter((u) => Boolean(u.paraId)), identifiedText);
  const rightIds = grouped(after.filter((u) => Boolean(u.paraId)), identifiedText);
  const repeats: Pair[] = [];
  for (const a of before) {
    if (usedBefore.has(a)) continue;
    const candidates = new Set<Unit>();
    const id = identifiedText(a);
    const right = rightIds.get(id);
    const sameIdTarget = leftIds.get(id)?.length === 1 && right?.length === 1 ? right[0] : undefined;
    if (sameIdTarget) candidates.add(sameIdTarget);
    for (const offset of [-1, 1]) {
      const matched = pairs.get(before[a.order + offset]);
      const candidate = matched && after[matched.to.order - offset];
      if (candidate?.key === a.key) candidates.add(candidate);
    }
    for (const b of candidates) {
      if (usedAfter.has(b)) continue;
      const sameId = b === sameIdTarget;
      const adjacent = neighbors(a, b);
      if (!sameId && !adjacent) continue;
      repeats.push({
        from: a, to: b, score: Number(sameId) * 3 + adjacent,
        evidence: ["exact", ...(sameId ? ["paragraphId" as const] : []), ...(adjacent ? ["neighbors" as const] : [])],
      });
    }
  }
  mutual(repeats, 1).forEach(pair);
  for (const b of after) {
    const sources = fromKeys.get(b.key);
    if (usedAfter.has(b) || sources?.length !== 1 || !pairs.has(sources[0])) continue;
    links.push(link("copied", sources, [b], ["exact"]));
    usedAfter.add(b);
  }
  const groups = groupMatches(before, after, usedBefore, usedAfter);
  for (const g of groups.matches) add(g.kind, g.from, g.to, ["joined"]);
  for (const g of groups.conflicts) add("ambiguous", g.from, g.to, ["joined"]);

  const index = new Map<string, Unit[]>();
  for (const b of after) {
    if (usedAfter.has(b) || fromKeys.has(b.key)) continue;
    for (const gram of b.grams) {
      const list = index.get(gram) ?? [];
      if (list.length <= 64) list.push(b);
      index.set(gram, list);
    }
  }
  const similar: Pair[] = [];
  for (const a of before) {
    if (usedBefore.has(a) || toKeys.has(a.key) || a.joinedKey.length < 8) continue;
    const hits = new Map<Unit, number>();
    for (const gram of a.grams) {
      const list = index.get(gram);
      if (!list || list.length > 64) continue;
      for (const b of list) hits.set(b, (hits.get(b) ?? 0) + 1);
    }
    const candidates = [...hits].sort((x, y) => y[1] - x[1]).slice(0, 24);
    for (const [b] of candidates) {
      if (a.kind !== b.kind || b.joinedKey.length < 8) continue;
      const chars = overlap(a.grams, b.grams);
      const sameId = Boolean(a.paraId && a.paraId === b.paraId);
      const adjacent = neighbors(a, b);
      const heading = Boolean(a.heading && a.heading === b.heading);
      const words = a.words && b.words ? overlap(a.words, b.words) : 0;
      if (chars < 0.4 || (chars < 0.62 && !sameId && !adjacent && !(words >= 0.65 && chars >= 0.5))) continue;
      const evidence: MatchEvidence[] = ["characterOverlap"];
      if (sameId) evidence.push("paragraphId");
      if (adjacent) evidence.push("neighbors");
      if (heading) evidence.push("heading");
      if (words >= 0.65) evidence.push("japaneseWords");
      similar.push({ from: a, to: b, score: chars + Number(sameId) * 0.13 + adjacent * 0.06 + Number(heading) * 0.035 + words * 0.04, evidence });
    }
  }
  mutual(similar, 0.075).filter((p) => p.score >= 0.62).forEach(pair);

  // Only an inversion among correspondences changes relative order; inserted paragraphs do not.
  const sorted = [...pairs].sort(([a], [b]) => a.order - b.order);
  let max = -1;
  for (const [, p] of sorted) {
    if (p.to.order < max) p.link.reordered = true;
    max = Math.max(max, p.to.order);
  }
  let min = Infinity;
  for (const [, p] of sorted.reverse()) {
    if (p.to.order > min) p.link.reordered = true;
    min = Math.min(min, p.to.order);
  }

  const similarByAfter = new Map<Unit, Pair[]>();
  for (const p of similar) {
    if (usedBefore.has(p.from) || usedAfter.has(p.to)) continue;
    const list = similarByAfter.get(p.to) ?? [];
    list.push(p);
    similarByAfter.set(p.to, list);
  }
  const possibleSources = new Set<Unit>();
  const ambiguousKeys = new Set<string>();
  function ambiguous(sources: Unit[], target: Unit, evidence: MatchEvidence[]) {
    const relationship = link("ambiguous", sources.slice(0, 12), [target], evidence);
    if (sources.length > relationship.from.length) relationship.candidateCount = sources.length;
    links.push(relationship);
  }
  for (const b of after) {
    if (usedAfter.has(b)) continue;
    const exact = fromKeys.get(b.key);
    if (exact?.length) {
      ambiguousKeys.add(b.key);
      ambiguous(exact, b, ["exact"]);
      continue;
    }
    const proposals = similarByAfter.get(b) ?? [];
    const options = proposals.map((p) => p.from);
    if (options.length) {
      options.forEach((a) => possibleSources.add(a));
      ambiguous(options, b, ["characterOverlap"]);
      continue;
    }
    const old = past.get(b.key);
    if (!fromKeys.has(b.key) && old?.length === 1 && toKeys.get(b.key)?.length === 1) {
      links.push(link("reappeared", old, [b], ["exact", "past"]));
    } else if (old?.length) {
      ambiguous(old, b, ["exact", "past"]);
    } else {
      links.push(link("added", [], [b]));
    }
  }
  for (const a of before) {
    if (!usedBefore.has(a) && !possibleSources.has(a) && !ambiguousKeys.has(a.key)) links.push(link("removed", [a], []));
  }
  return links.sort((a, b) =>
    (a.to[0]?.paragraph ?? Infinity) - (b.to[0]?.paragraph ?? Infinity)
    || (a.to[0]?.start ?? Infinity) - (b.to[0]?.start ?? Infinity)
    || (a.from[0]?.paragraph ?? Infinity) - (b.from[0]?.paragraph ?? Infinity)
    || (a.from[0]?.start ?? Infinity) - (b.from[0]?.start ?? Infinity),
  );
}
