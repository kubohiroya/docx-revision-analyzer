import type { SavedParagraphRange, SavedTextVersion } from "./lineage";
import { prepareUnits, type Unit } from "./lineageMatching";

/** IDs distinguish versions within a comparison, including copies with identical bytes. */
export type VersionBranchInput = SavedTextVersion & {
  id: string;
  fileCreatedAt?: string;
  /** Only provide this when the versions came from the same known archive. */
  archiveId?: string;
};

type TimeSource = "archivedAt" | "sourceModified" | "fileCreatedAt";
type TimeUncertainty = "missing-time" | "creation-time-only" | "conflicting-times";

export interface VersionBranchTime {
  observed: Partial<Record<TimeSource, string>>;
  selected?: { source: TimeSource; value: string };
  archiveNumber?: number;
  uncertainty?: TimeUncertainty;
}

type CandidateOrder = "earlier" | "same-time" | "unconfirmed-time" | "archive-order-conflict";

export interface BranchSentenceMatch {
  from: SavedParagraphRange;
  to: SavedParagraphRange;
  /** This sentence occurs in only one potential source in the selected set. */
  distinctive: boolean;
}

export interface VersionBranchCandidate {
  parentId: string;
  order: CandidateOrder;
  matches: BranchSentenceMatch[];
  distinctiveMatches: number;
  /** Fractions of eligible body sentences, weighted by inverse document frequency. */
  sourceCoverage: number;
  targetCoverage: number;
}

export type VersionBranchHoldReason = TimeUncertainty | "unavailable" | "duplicate-file" | "same-visible-text"
  | "no-earlier-version" | "no-shared-body-text" | "only-common-body-text" | "time-order-unconfirmed" | "competing-candidates"
  | "insufficient-distinctive-text" | "insufficient-coverage";

export interface VersionBranchEntry {
  id: string;
  time: VersionBranchTime;
  duplicateOf?: string;
  sameVisibleText: string[];
  candidates: VersionBranchCandidate[];
  decision: { status: "inferred"; parentId: string } | { status: "held"; reason: VersionBranchHoldReason };
}

interface PreparedVersion {
  input: VersionBranchInput;
  time: VersionBranchTime;
  epoch?: number;
  sentences: Map<string, Unit | null>;
  signature?: string;
}

const sentenceSegmenter = new Intl.Segmenter("ja", { granularity: "sentence" });
const minimumSentenceLength = 12;
const minimumDistinctiveMatches = 2;
const minimumSourceCoverage = 0.5;
const minimumTargetCoverage = 0.3;

function compareIds(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0;
}

function parseTime(value?: string): number | undefined {
  // Readers supply ISO timestamps. Local times without an offset cannot establish a cross-file order.
  if (!value || !/T.*(?:Z|[+-]\d{2}:?\d{2})$/iu.test(value)) return;
  const timestamp = Date.parse(value);
  return Number.isFinite(timestamp) ? timestamp : undefined;
}

function prepareVersion(input: VersionBranchInput): PreparedVersion {
  const observed: VersionBranchTime["observed"] = {};
  for (const source of ["archivedAt", "sourceModified", "fileCreatedAt"] as const) {
    const value = input[source];
    if (value !== undefined) observed[source] = value;
  }
  const archived = input.source === "archive" ? parseTime(observed.archivedAt) : undefined;
  const modified = parseTime(observed.sourceModified);
  const created = parseTime(observed.fileCreatedAt);
  const epoch = archived ?? modified ?? created;
  const time: VersionBranchTime = { observed };
  if (input.source === "archive" && input.number !== undefined) time.archiveNumber = input.number;
  if (epoch === undefined) {
    time.uncertainty = "missing-time";
  } else {
    const source = archived !== undefined ? "archivedAt" : modified !== undefined ? "sourceModified" : "fileCreatedAt";
    time.selected = { source, value: new Date(epoch).toISOString() };
    if ((created !== undefined && modified !== undefined && created > modified)
      || (archived !== undefined && modified !== undefined && modified > archived)
      || (archived !== undefined && created !== undefined && created > archived)) {
      time.uncertainty = "conflicting-times";
    } else if (source === "fileCreatedAt") {
      time.uncertainty = "creation-time-only";
    }
  }

  const units = input.unavailable ? [] : prepareUnits(input.paragraphs, 0, sentenceSegmenter);
  const sentences = new Map<string, Unit | null>();
  for (const unit of units) {
    if (unit.kind !== "body" || Array.from(unit.joinedKey).length < minimumSentenceLength) continue;
    // A repeated sentence affects specificity but cannot locate a unique match.
    sentences.set(unit.key, sentences.has(unit.key) ? null : unit);
  }
  const signature = units.length ? JSON.stringify(units.map(({ kind, key }) => [kind, key])) : undefined;
  return { input, time, epoch, sentences, signature };
}

function potentialOrder(parent: PreparedVersion, child: PreparedVersion): CandidateOrder | "later" {
  if (parent.epoch === undefined || child.epoch === undefined || parent.time.uncertainty || child.time.uncertainty) {
    return "unconfirmed-time";
  }
  if (parent.epoch === child.epoch) return "same-time";
  const left = parent.time.archiveNumber;
  const right = child.time.archiveNumber;
  if (parent.input.archiveId && parent.input.archiveId === child.input.archiveId && left !== undefined && right !== undefined
    && (left === right || (left < right) !== (parent.epoch < child.epoch))) return "archive-order-conflict";
  return parent.epoch < child.epoch ? "earlier" : "later";
}

function paragraphRange(unit: Unit): SavedParagraphRange {
  const { paragraph, start, end } = unit.range;
  return { paragraph, start, end };
}

function collectCandidates(
  child: PreparedVersion, bySentence: ReadonlyMap<string, PreparedVersion[]>, totals: ReadonlyMap<PreparedVersion, number>,
): VersionBranchCandidate[] {
  const grouped = new Map<PreparedVersion, { order: CandidateOrder; matches: BranchSentenceMatch[]; weight: number }>();
  for (const [key, to] of child.sentences) {
    const owners = bySentence.get(key);
    if (!to || !owners) continue;
    const sources = owners.flatMap(parent => {
      if (parent === child) return [];
      const order = potentialOrder(parent, child);
      return order === "later" ? [] : [{ parent, order }];
    });
    for (const { parent, order } of sources) {
      const from = parent.sentences.get(key);
      if (!from) continue;
      const item = grouped.get(parent) ?? { order, matches: [], weight: 0 };
      item.matches.push({ from: paragraphRange(from), to: paragraphRange(to), distinctive: sources.length === 1 });
      item.weight += 1 / owners.length;
      grouped.set(parent, item);
    }
  }

  return [...grouped].map(([parent, { order, matches, weight }]): VersionBranchCandidate => ({
    parentId: parent.input.id, order, matches, distinctiveMatches: matches.filter(match => match.distinctive).length,
    sourceCoverage: weight / (totals.get(parent) || 1), targetCoverage: weight / (totals.get(child) || 1),
  })).sort((a, b) => Number(b.order === "earlier") - Number(a.order === "earlier")
    || b.distinctiveMatches - a.distinctiveMatches || b.sourceCoverage - a.sourceCoverage || b.targetCoverage - a.targetCoverage
    || compareIds(a.parentId, b.parentId));
}

function decideParent(
  version: PreparedVersion, duplicateOf: string | undefined, sameVisibleText: string[],
  candidates: VersionBranchCandidate[], comparable: PreparedVersion[], commonSentences: ReadonlySet<string>,
): VersionBranchEntry["decision"] {
  const hold = (reason: VersionBranchHoldReason) => ({ status: "held", reason } as const);
  if (version.input.unavailable) return hold("unavailable");
  if (duplicateOf !== undefined) return hold("duplicate-file");
  if (version.time.uncertainty) return hold(version.time.uncertainty);
  if (sameVisibleText.length) return hold("same-visible-text");

  const best = candidates.find(candidate => candidate.order === "earlier");
  if (!best) {
    if (candidates.length) return hold("time-order-unconfirmed");
    if (!comparable.some(parent => parent !== version && potentialOrder(parent, version) === "earlier")) {
      return hold("no-earlier-version");
    }
    for (const [key, sentence] of version.sentences) {
      if (sentence && commonSentences.has(key)) return hold("only-common-body-text");
    }
    return hold("no-shared-body-text");
  }
  if (candidates.some(candidate => candidate !== best && candidate.distinctiveMatches > 0)) return hold("competing-candidates");
  if (best.distinctiveMatches < minimumDistinctiveMatches) return hold("insufficient-distinctive-text");
  if (best.sourceCoverage < minimumSourceCoverage || best.targetCoverage < minimumTargetCoverage) return hold("insufficient-coverage");
  return { status: "inferred", parentId: best.parentId };
}

/** Infer possible parents within a set of saved versions and retain the evidence for review. */
export function inferVersionBranches(input: readonly VersionBranchInput[]): VersionBranchEntry[] {
  if (new Set(input.map(version => version.id)).size !== input.length) throw new Error("Version IDs must be unique");
  const compareTimes = (a: PreparedVersion, b: PreparedVersion) =>
    (a.epoch ?? Infinity) - (b.epoch ?? Infinity) || compareIds(a.input.id, b.input.id);
  const versions = input.map(prepareVersion).sort(compareTimes);
  const representatives = new Map<string, string>();
  const duplicates = new Map<PreparedVersion, string>();
  const readable = versions.filter(version => !version.input.unavailable);
  // Prefer a reliable timestamp when choosing a representative, then the earliest observation.
  for (const version of [...readable].sort((a, b) => Number(Boolean(a.time.uncertainty)) - Number(Boolean(b.time.uncertainty))
    || compareTimes(a, b))) {
    const hash = version.input.sha256?.toLowerCase();
    if (!hash) continue;
    const representative = representatives.get(hash);
    if (representative === undefined) representatives.set(hash, version.input.id);
    else duplicates.set(version, representative);
  }
  const comparable = readable.filter(version => !duplicates.has(version));
  const bySentence = new Map<string, PreparedVersion[]>();
  const byText = new Map<string, PreparedVersion[]>();
  for (const version of comparable) {
    if (version.signature !== undefined) {
      const matches = byText.get(version.signature) ?? [];
      matches.push(version);
      byText.set(version.signature, matches);
    }
    for (const key of version.sentences.keys()) {
      const owners = bySentence.get(key) ?? [];
      owners.push(version);
      bySentence.set(key, owners);
    }
  }
  const commonSentences = new Set<string>();
  if (comparable.length > 2) {
    // Corpus-wide sentences cannot distinguish the branches in this set.
    for (const [key, owners] of bySentence) {
      if (owners.length !== comparable.length) continue;
      commonSentences.add(key);
      bySentence.delete(key);
    }
  }
  const totals = new Map(comparable.map(version => {
    let total = 0;
    for (const [key, sentence] of version.sentences) {
      const owners = bySentence.get(key);
      if (sentence && owners) total += 1 / owners.length;
    }
    return [version, total] as const;
  }));

  return versions.map(version => {
    const duplicateOf = duplicates.get(version);
    const comparableVersion = !version.input.unavailable && duplicateOf === undefined;
    const matches = comparableVersion && version.signature !== undefined ? byText.get(version.signature) ?? [] : [];
    const sameVisibleText = matches.filter(parent => parent !== version && potentialOrder(parent, version) !== "later")
      .map(parent => parent.input.id);
    const candidates = comparableVersion ? collectCandidates(version, bySentence, totals) : [];
    return {
      id: version.input.id, time: version.time, ...(duplicateOf === undefined ? {} : { duplicateOf }), sameVisibleText, candidates,
      decision: decideParent(version, duplicateOf, sameVisibleText, candidates, comparable, commonSentences),
    };
  });
}
