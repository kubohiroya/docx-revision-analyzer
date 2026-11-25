import { createHash } from "node:crypto";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { readSavedTextVersion, type SavedTextVersion, type VersionIdentity } from "../../src/lib/lineage";
import type { WatchArchive } from "../../src/node/docWatcher";
import { readArchivedSnapshot } from "../../src/node/snapshotArchive";

export async function readSelectedSavedVersions(files: string[]): Promise<SavedTextVersion[]> {
  const result: SavedTextVersion[] = [];
  for (const file of files) {
    const identity: VersionIdentity = { label: path.basename(file), source: "selected" };
    try {
      const [bytes, stat] = await Promise.all([fs.readFile(file), fs.stat(file)]);
      result.push(await readSavedTextVersion(bytes, {
        ...identity,
        sha256: createHash("sha256").update(bytes).digest("hex"),
        sourceModified: stat.mtime.toISOString(),
      }));
    } catch {
      result.push({ ...identity, unavailable: true });
    }
  }
  return result;
}

export async function readSavedVersionArchive(archive: WatchArchive): Promise<SavedTextVersion[]> {
  const versions = archive.versions().slice().sort((a, b) => a.n - b.n);
  const resumed = new Set<number>();
  const latestFirst = [...versions].reverse();
  for (const gap of archive.info.gaps ?? []) {
    const match = latestFirst.find((v) => v.sha256 === gap.sha256 && v.archivedAt <= gap.detectedAt);
    if (match) resumed.add(match.n);
  }
  const result: SavedTextVersion[] = [];
  for (const [index, entry] of versions.entries()) {
    const previous = versions[index - 1];
    const identity: VersionIdentity = {
      label: entry.file, source: "archive", number: entry.n, sha256: entry.sha256,
      archivedAt: entry.archivedAt, sourceModified: entry.sourceModified,
      gapBefore: resumed.has(entry.n) || Boolean(previous && entry.n !== previous.n + 1),
    };
    try {
      result.push(await readSavedTextVersion(await readArchivedSnapshot(archive.dir, entry), identity));
    } catch {
      result.push({ ...identity, unavailable: true });
    }
  }
  return result;
}
