import { createHash } from "node:crypto";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { extractComments, replayCommentEvents, type CommentEvent, type CommentLocations, type DocxComment } from "../../src/lib/docxComments";
import { readSavedTextVersion, type SavedParagraphRange, type VersionIdentity } from "../../src/lib/lineage";
import type { WatchArchive } from "../../src/node/docWatcher";
import { readArchivedSnapshot, type SnapshotEntry } from "../../src/node/snapshotArchive";
import { commentObservations, commentThreads, observedVersion } from "./commentHistory";
import type { CommentObservation, SavedLineageVersion, SavedVersionThread } from "./shared";

function makeThreads(
  comments: DocxComment[],
  identity: VersionIdentity,
  observations: ReadonlyMap<string, CommentObservation[]>,
  locations?: CommentLocations,
  ranges?: Map<string, SavedParagraphRange[]>,
): SavedVersionThread[] {
  const grouped = commentThreads(comments);
  const uniqueIds = new Set<string>();
  const duplicateIds = new Set<string>();
  for (const comment of comments) {
    if (uniqueIds.has(comment.id)) duplicateIds.add(comment.id);
    uniqueIds.add(comment.id);
  }
  const rootKeys = new Set<string>();
  const duplicateKeys = new Set<string>();
  for (const [root] of grouped) {
    if (rootKeys.has(root.key)) duplicateKeys.add(root.key);
    rootKeys.add(root.key);
  }
  return grouped.map((thread) => {
    const root = thread[0];
    const unique = !duplicateIds.has(root.id);
    return {
      key: root.key,
      ranges: unique ? ranges?.get(root.id) ?? [] : [],
      detail: {
        version: identity.number !== undefined && identity.archivedAt
          ? { n: identity.number, archivedAt: identity.archivedAt }
          : undefined,
        versionUnavailable: !locations,
        passage: unique ? locations?.passages.get(root.id) : undefined,
        recordedAnchor: root.anchor,
        comments: thread,
        observations: duplicateKeys.has(root.key) ? [] : observations.get(root.key) ?? [],
      },
    };
  });
}

async function readVersion(
  bytes: Uint8Array, identity: VersionIdentity, observations: ReadonlyMap<string, CommentObservation[]> = new Map(),
): Promise<SavedLineageVersion> {
  const ranges = new Map<string, SavedParagraphRange[]>();
  const version = await readSavedTextVersion(bytes, identity, ranges);
  const locations: CommentLocations = { passages: new Map() };
  try {
    const comments = await extractComments(bytes, locations);
    return { ...version, threads: makeThreads(comments, identity, observations, locations, ranges) };
  } catch {
    return { ...version, threads: [], commentsUnavailable: true };
  }
}

export async function readSelectedSavedVersions(files: string[]): Promise<SavedLineageVersion[]> {
  const result: SavedLineageVersion[] = [];
  for (const file of files) {
    const identity: VersionIdentity = { label: path.basename(file), source: "selected" };
    try {
      const [bytes, stat] = await Promise.all([fs.readFile(file), fs.stat(file)]);
      result.push(await readVersion(bytes, {
        ...identity,
        sha256: createHash("sha256").update(bytes).digest("hex"),
        sourceModified: stat.mtime.toISOString(),
      }));
    } catch {
      result.push({ ...identity, unavailable: true, threads: [] });
    }
  }
  return result;
}

export async function readSavedVersionArchive(archive: WatchArchive): Promise<SavedLineageVersion[]> {
  const versions = archive.versions().slice().sort((a, b) => a.n - b.n);
  const events = archive.readEvents();
  const observations = commentObservations(events, versions);
  const resumed = new Set<number>();
  const latestFirst = [...versions].reverse();
  for (const gap of archive.info.gaps ?? []) {
    const match = latestFirst.find((v) => v.sha256 === gap.sha256 && v.archivedAt <= gap.detectedAt);
    if (match) resumed.add(match.n);
  }
  const result: SavedLineageVersion[] = [];
  for (const [index, entry] of versions.entries()) {
    const previous = versions[index - 1];
    const identity: VersionIdentity = {
      label: entry.file, source: "archive", number: entry.n, sha256: entry.sha256,
      archivedAt: entry.archivedAt, sourceModified: entry.sourceModified,
      gapBefore: resumed.has(entry.n) || Boolean(previous && entry.n !== previous.n + 1),
    };
    try {
      result.push(await readVersion(await readArchivedSnapshot(archive.dir, entry), identity, observations));
    } catch {
      const comments = recordedCommentsAt(events, versions, entry.n);
      result.push({ ...identity, unavailable: true, threads: makeThreads(comments, identity, observations) });
    }
  }
  return result;
}

/** Use the local log as a fallback only when an observation identifies this archived version. */
function recordedCommentsAt(events: CommentEvent[], versions: SnapshotEntry[], number: number): DocxComment[] {
  let last = events.length - 1;
  while (last >= 0 && (events[last].type === "receipt" || observedVersion(events[last], versions)?.n !== number)) last--;
  if (last < 0) return [];
  return [...replayCommentEvents(events.slice(0, last + 1)).values()]
    .filter((state) => state.present).map((state) => state.comment);
}
