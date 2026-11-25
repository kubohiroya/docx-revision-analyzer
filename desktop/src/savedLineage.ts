import { createHash, randomUUID } from "node:crypto";
import type { Stats } from "node:fs";
import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import * as path from "node:path";
import { addWordComment, extractComments, replayCommentEvents, type CommentEvent, type CommentLocations, type DocxComment } from "../../src/lib/docxComments";
import { readSavedTextVersion, type SavedParagraphRange, type VersionIdentity } from "../../src/lib/lineage";
import type { WatchArchive } from "../../src/node/docWatcher";
import { checkNotOpenElsewhere, replaceWithBackup } from "../../src/node/historyFile";
import { readArchivedSnapshot, type SnapshotEntry } from "../../src/node/snapshotArchive";
import { commentObservations, commentThreads, observedVersion } from "./commentHistory";
import type { CommentObservation, SavedLineageVersion, SavedVersionCommentError, SavedVersionCommentRequest,
  SavedVersionCommentResult, SavedVersionThread } from "./shared";

const versionFiles = new Map<string, () => Promise<string>>();
const localVersions = new Map<string, { file: string; sha: string; dev: number; ino: number }>();
const pendingCommentWrites = new Map<string, Promise<void>>();

type LocalVersion = NonNullable<ReturnType<typeof localVersions.get>>;

function digest(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function rememberVersionFile(resolve: () => Promise<string>): string {
  const id = randomUUID();
  versionFiles.set(id, resolve);
  return id;
}

export async function savedVersionFile(id: unknown): Promise<string> {
  const resolve = typeof id === "string" ? versionFiles.get(id) : undefined;
  if (!resolve) throw new Error("The saved version is no longer available.");
  return resolve();
}

function rememberLocalVersion(file: string, stat: Stats, sha: string): string {
  const source = { file, sha, dev: stat.dev, ino: stat.ino };
  const id = rememberVersionFile(async () => {
    const actual = await fs.lstat(file);
    if (!actual.isFile() || actual.dev !== stat.dev || actual.ino !== stat.ino || digest(await fs.readFile(file)) !== sha) {
      throw new Error("The DOCX has changed since it was compared. Load the versions again.");
    }
    return file;
  });
  localVersions.set(id, source);
  return id;
}

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
    return { ...version, fileSize: bytes.byteLength, threads: makeThreads(comments, identity, observations, locations, ranges) };
  } catch {
    return { ...version, fileSize: bytes.byteLength, threads: [], commentsUnavailable: true };
  }
}

export async function readSelectedSavedVersions(files: string[], allowOpen = false): Promise<SavedLineageVersion[]> {
  const result: SavedLineageVersion[] = [];
  for (const file of files) {
    const stat = await fs.stat(file).catch(() => undefined);
    const identity: VersionIdentity = {
      label: path.basename(file), source: "selected", sourceModified: stat?.mtime.toISOString(),
    };
    const fileCreatedAt = stat && stat.birthtimeMs > 0 ? stat.birthtime.toISOString() : undefined;
    try {
      const actual = await fs.realpath(file);
      const bytes = await fs.readFile(actual);
      const sha256 = digest(bytes);
      const version = await readVersion(bytes, { ...identity, sha256 });
      const openId = allowOpen && stat?.isFile() && /\.docx$/iu.test(actual) ? rememberLocalVersion(actual, stat, sha256) : undefined;
      result.push({ ...version, fileCreatedAt, openId });
    } catch {
      result.push({ ...identity, fileCreatedAt, fileSize: stat?.size, unavailable: true, threads: [] });
    }
  }
  return result;
}

class CommentWriteConflict extends Error {
  constructor(readonly reason: SavedVersionCommentError) { super(reason); }
}

async function currentWritableVersion(source: LocalVersion): Promise<Uint8Array> {
  const { file, sha, dev, ino } = source;
  const current = await fs.lstat(file).catch(() => undefined);
  if (!current?.isFile() || current.dev !== dev || current.ino !== ino) throw new CommentWriteConflict("changed");
  if (!(current.mode & 0o222)) throw new CommentWriteConflict("readOnly");
  if (checkNotOpenElsewhere(file)) throw new CommentWriteConflict("locked");
  if (!await fs.access(file, fs.constants.W_OK).then(() => true, () => false)) throw new CommentWriteConflict("readOnly");
  const bytes = await fs.readFile(file).catch(() => undefined);
  if (!bytes || digest(bytes) !== sha) throw new CommentWriteConflict("changed");
  const check = await fs.lstat(file).catch(() => undefined);
  if (!check?.isFile() || check.dev !== dev || check.ino !== ino || check.mtimeMs !== current.mtimeMs ||
      check.ctimeMs !== current.ctimeMs) throw new CommentWriteConflict("changed");
  return bytes;
}

function validComment(value: unknown): value is SavedVersionCommentRequest {
  const comment = value as Partial<SavedVersionCommentRequest> | undefined;
  const range = comment?.range;
  return typeof comment?.author === "string" && typeof comment.text === "string" && typeof range?.paragraphText === "string" &&
    (range.paragraphId === undefined || typeof range.paragraphId === "string") &&
    Number.isSafeInteger(range.start) && Number.isSafeInteger(range.end);
}

export async function addSavedVersionComment(id: unknown, request: unknown): Promise<SavedVersionCommentResult> {
  const source = typeof id === "string" ? localVersions.get(id) : undefined;
  if (!source) return { ok: false, reason: "unavailable" };
  if (!validComment(request)) return { ok: false, reason: "unsupported" };
  const { file } = source;
  const previous = pendingCommentWrites.get(file) ?? Promise.resolve();
  const result = previous.then(async (): Promise<SavedVersionCommentResult> => {
    try {
      const bytes = await currentWritableVersion(source);
      const label = path.basename(file);
      const generated = await addWordComment(bytes, request).catch(() => { throw new CommentWriteConflict("unsupported"); });
      const sha256 = digest(generated.bytes);
      const version = await readVersion(generated.bytes, { label, source: "selected", sha256 });
      if (version.unavailable || version.commentsUnavailable ||
          !version.threads.some((thread) => thread.detail.comments[0]?.id === generated.id && thread.ranges.length)) {
        return { ok: false, reason: "unsupported" };
      }
      await currentWritableVersion(source);
      const backupPath = await replaceWithBackup(file, generated.bytes, async (backup) => {
        if (digest(await fs.readFile(backup)) !== source.sha) throw new CommentWriteConflict("changed");
        await currentWritableVersion(source);
      });
      const stat = await fs.stat(file).catch(() => undefined);
      return { ok: true, backupPath, version: { ...version,
        sourceModified: stat?.mtime.toISOString(), fileCreatedAt: stat && stat.birthtimeMs > 0 ? stat.birthtime.toISOString() : undefined,
        openId: stat?.isFile() ? rememberLocalVersion(file, stat, sha256) : undefined,
      } };
    } catch (error) {
      return { ok: false, reason: error instanceof CommentWriteConflict ? error.reason : "saveFailed" };
    }
  });
  const settled = result.then(() => undefined, () => undefined);
  pendingCommentWrites.set(file, settled);
  try {
    return await result;
  } finally {
    if (pendingCommentWrites.get(file) === settled) pendingCommentWrites.delete(file);
  }
}

export async function readSavedVersionArchive(archive: WatchArchive, allowOpen = false): Promise<SavedLineageVersion[]> {
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
      const version = await readVersion(await readArchivedSnapshot(archive.dir, entry), identity, observations);
      const openId = allowOpen ? rememberVersionFile(async () => {
        const bytes = await readArchivedSnapshot(archive.dir, entry);
        const folder = await fs.mkdtemp(path.join(tmpdir(), "dra-saved-version-"));
        const copy = path.join(folder, entry.file);
        try {
          await fs.writeFile(copy, bytes, { flag: "wx", mode: 0o400 });
          return copy;
        } catch (error) {
          await fs.rm(folder, { recursive: true, force: true });
          throw error;
        }
      }) : undefined;
      result.push({ ...version, openId });
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
