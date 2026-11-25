import { commentReferences } from "../../src/lib/commentReferences";
import {
  extractComments,
  isReceipt,
  replayCommentEvents,
  type CommentEvent,
  type CommentLocations,
  type DocxComment,
} from "../../src/lib/docxComments";
import type { WatchArchive } from "../../src/node/docWatcher";
import { readArchivedSnapshot, type SnapshotEntry } from "../../src/node/snapshotArchive";
import type { CommentHistoryDetail, CommentHistoryRow, CommentObservation } from "./shared";

function rootKey(comment: DocxComment, byKey: Map<string, DocxComment>): string {
  const seen = new Set([comment.key]);
  while (comment.parentKey) {
    const parent = byKey.get(comment.parentKey);
    if (!parent) break;
    if (seen.has(parent.key)) return [...seen].sort()[0];
    seen.add(parent.key);
    comment = parent;
  }
  return comment.key;
}

export function listCommentHistory(archive: WatchArchive): CommentHistoryRow[] {
  const events = archive.readEvents();
  const states = replayCommentEvents(events);
  const byKey = new Map([...states].map(([key, state]) => [key, state.comment]));
  const membersByRoot = new Map<string, string[]>();
  for (const comment of byKey.values()) {
    const root = rootKey(comment, byKey);
    const members = membersByRoot.get(root) ?? [];
    members.push(comment.key);
    membersByRoot.set(root, members);
  }
  const tagsByRoot = new Map<string, Set<string>>();
  for (const { comment } of events) {
    if (!comment || isReceipt(comment)) continue;
    const root = rootKey(comment, byKey);
    const tags = tagsByRoot.get(root) ?? new Set<string>();
    for (const tag of commentReferences(comment.text, comment.links).tags) tags.add(tag);
    tagsByRoot.set(root, tags);
  }
  return [...states.values()]
    .filter(({ comment }) => !isReceipt(comment) && rootKey(comment, byKey) === comment.key)
    .map((state) => ({ ...state, members: membersByRoot.get(state.key) ?? [], tags: [...(tagsByRoot.get(state.key) ?? [])] }));
}

export function observedVersion(event: CommentEvent, versions: SnapshotEntry[]): SnapshotEntry | undefined {
  if (!event.sha256) return undefined;
  const matching = versions.filter((version) => version.sha256 === event.sha256);
  const atObservation = matching.filter((version) => version.archivedAt === event.at);
  if (atObservation.length === 1) return atObservation[0];
  return matching.length === 1 ? matching[0] : undefined;
}

export function commentObservations(events: CommentEvent[], versions: SnapshotEntry[]): Map<string, CommentObservation[]> {
  const result = new Map<string, CommentObservation[]>();
  for (const event of events) {
    const { type, key, at } = event;
    if (type !== "resolved" && type !== "reopened" && type !== "removed" && type !== "reappeared") continue;
    const thread = result.get(key) ?? [];
    thread.push({ type, at, version: observedVersion(event, versions)?.n });
    result.set(key, thread);
  }
  return result;
}

/** Group replies from a single version. Reused legacy keys cannot identify a unique parent. */
export function commentThreads(comments: DocxComment[]): DocxComment[][] {
  const visible = comments.filter((comment) => !isReceipt(comment));
  const byKey = new Map<string, DocxComment>();
  const ambiguous = new Set<string>();
  for (const comment of visible) {
    if (byKey.has(comment.key)) ambiguous.add(comment.key);
    byKey.set(comment.key, comment);
  }
  for (const key of ambiguous) byKey.delete(key);
  const groups = new Map<DocxComment, DocxComment[]>();
  for (const comment of visible) {
    const root = byKey.get(rootKey(comment, byKey)) ?? comment;
    const thread = groups.get(root) ?? [root];
    if (root !== comment) thread.push(comment);
    groups.set(root, thread);
  }
  return [...groups.values()];
}

export async function readCommentHistory(archive: WatchArchive, key: string): Promise<CommentHistoryDetail | undefined> {
  const events = archive.readEvents();
  const states = replayCommentEvents(events);
  const state = states.get(key);
  if (!state) return undefined;

  let versions: SnapshotEntry[];
  try {
    versions = archive.versions();
  } catch {
    versions = [];
  }
  const transitions = events.filter((event): event is CommentEvent & { type: CommentObservation["type"] } =>
    event.key === key && ["resolved", "reopened", "removed", "reappeared"].includes(event.type),
  );
  const observations = commentObservations(transitions, versions).get(key) ?? [];
  const lastRemoval = state.present ? undefined : transitions.filter((e) => e.type === "removed").at(-1);
  const absentVersion = lastRemoval && observedVersion(lastRemoval, versions);
  const candidates = absentVersion ? versions.filter((v) => v.n < absentVersion.n) : versions;
  let versionUnavailable = false;

  for (const version of [...candidates].reverse()) {
    const locations: CommentLocations = { passages: new Map() };
    let comments: DocxComment[];
    try {
      comments = await extractComments(await readArchivedSnapshot(archive.dir, version), locations);
    } catch {
      versionUnavailable = true;
      continue;
    }
    const roots = comments.filter((c) => c.key === key && !isReceipt(c));
    if (roots.length !== 1) continue;
    const root = roots[0];
    const exact = locations.passages.get(root.id);
    const recordedAnchor = root.anchor ?? events
      .filter((event) => event.key === key && event.comment?.anchor && event.at <= version.archivedAt)
      .at(-1)?.comment?.anchor;
    const inferred = !exact && recordedAnchor && locations.find?.(recordedAnchor);
    const byKey = new Map(comments.map((comment) => [comment.key, comment]));
    const replies = comments.filter((c) => c.key !== key && rootKey(c, byKey) === key && !isReceipt(c));
    return {
      version: { n: version.n, archivedAt: version.archivedAt },
      versionUnavailable,
      passage: exact ?? (inferred || undefined),
      locatedByContext: Boolean(inferred),
      recordedAnchor,
      comments: [root, ...replies],
      observations,
    };
  }

  // When the DOCX is missing, replay to just before the observation that lost the root.
  let cutoff = lastRemoval ? events.indexOf(lastRemoval) : events.length;
  while (lastRemoval && cutoff > 0 && events[cutoff - 1].at === lastRemoval.at && events[cutoff - 1].sha256 === lastRemoval.sha256) {
    cutoff--;
  }
  const recorded = lastRemoval ? replayCommentEvents(events.slice(0, cutoff)) : states;
  const root = recorded.get(key)?.comment ?? state.comment;
  const byKey = new Map([...recorded].filter(([, entry]) => entry.present).map(([id, entry]) => [id, entry.comment]));
  const replies = [...byKey.values()].filter((c) => c.key !== key && rootKey(c, byKey) === key && !isReceipt(c));
  return { versionUnavailable: true, recordedAnchor: root.anchor, comments: [root, ...replies], observations };
}
