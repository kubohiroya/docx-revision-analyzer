import { createHash, randomUUID } from "node:crypto";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { parseDocxLayout } from "../../src/lib/docxLayout";
import type { VersionIdentity } from "../../src/lib/lineage";
import { CloudVersionError, type CloudVersionItem } from "./cloudVersionGraph";
import type { MicrosoftClient } from "./microsoft";
import { readCloudSavedVersion } from "./savedLineage";
import type { SavedLineageVersion } from "./shared";

export interface CloudVersionTarget {
  driveId: string;
  itemId: string;
  name: string;
  account: string;
  mtime: string;
  bytes: Uint8Array;
}

interface ManifestEntry { item: CloudVersionItem; sha256: string; fetchedAt: string; order: number }

export function hash(input: string | Uint8Array): string {
  return createHash("sha256").update(input).digest("hex");
}

async function readManifest(file: string): Promise<ManifestEntry | undefined> {
  try {
    const entry = JSON.parse(await fs.readFile(file, "utf8")) as ManifestEntry;
    if (!/^[0-9a-f]{64}$/.test(entry.sha256) || typeof entry.item?.id !== "string" ||
        (entry.item.modified !== undefined && typeof entry.item.modified !== "string") || typeof entry.fetchedAt !== "string") {
      throw new CloudVersionError("invalid");
    }
    return entry;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
    throw error;
  }
}

async function replace(file: string, content: string | Uint8Array): Promise<void> {
  const temporary = `${file}.${randomUUID()}.tmp`;
  try {
    await fs.writeFile(temporary, content, { flag: "wx", mode: 0o600 });
    await fs.rename(temporary, file);
  } finally {
    await fs.rm(temporary, { force: true });
  }
}

export type CloudVersionList = { items: CloudVersionItem[]; error?: string };

export type CloudVersionImportResult = { versions: SavedLineageVersion[]; failures: { id: string; reason: string }[] };

export class CloudVersionImport {
  private controller?: AbortController;
  private newestFirst: CloudVersionItem[] = [];

  constructor(
    private readonly root: string,
    private readonly target: CloudVersionTarget,
    private readonly client: MicrosoftClient,
  ) {}

  cancel(): void { this.controller?.abort(); }

  private async verifyAccount(): Promise<void> {
    if ((await this.client.account())?.homeAccountId !== this.target.account) throw new CloudVersionError("sign-in");
  }

  private get documentId(): string {
    return hash(JSON.stringify([this.target.account, this.target.driveId, this.target.itemId]));
  }

  private async listCached(): Promise<CloudVersionItem[]> {
    const dir = path.join(this.root, this.documentId);
    const names = await fs.readdir(dir).catch(error => {
      if (error.code === "ENOENT") return [];
      throw error;
    });
    const entries: ManifestEntry[] = [];
    for (const name of names) {
      if (!/^[0-9a-f]{64}\.json$/.test(name)) continue;
      const entry = await readManifest(path.join(dir, name)).catch(() => undefined);
      if (entry && name === `${hash(entry.item.id)}.json`) entries.push(entry);
    }
    entries.sort((a, b) => (b.item.modified ?? "").localeCompare(a.item.modified ?? "") || b.order - a.order);
    return entries.map(entry => ({ ...entry.item, cachedOnly: true }));
  }

  async list(): Promise<CloudVersionList> {
    await this.verifyAccount();
    const signal = AbortSignal.timeout(60_000);
    let fresh: CloudVersionItem[] = [];
    let error: string | undefined;
    try {
      fresh = await this.client.listVersions(this.target, this.target.account, signal);
    } catch (cause) {
      error = cause instanceof CloudVersionError ? cause.reason : "network";
      // Permission failures must not be mistaken for a known gap in the server history.
      if (error === "sign-in" || error === "forbidden") {
        this.newestFirst = [];
        throw cause;
      }
    }
    const freshIds = new Set(fresh.map(item => item.id));
    const cached = (await this.listCached()).filter(item => !freshIds.has(item.id));
    await this.verifyAccount();
    this.newestFirst = [...fresh, ...cached];
    return { items: this.newestFirst, error };
  }

  async import(
    ids: string[], progress: (done: number, total: number) => void,
  ): Promise<CloudVersionImportResult> {
    if (this.controller) throw new CloudVersionError("conflict");
    const controller = new AbortController();
    this.controller = controller;
    try {
      await this.verifyAccount();
      const signal = controller.signal;
      const selected = new Set(ids);
      if (!selected.size || ids.some(id => !this.newestFirst.some(item => item.id === id))) throw new CloudVersionError("invalid");
      const document = this.documentId;
      const dir = path.join(this.root, document);
      await fs.mkdir(dir, { recursive: true, mode: 0o700 });
      const versions: SavedLineageVersion[] = [];
      const failures: CloudVersionImportResult["failures"] = [];
      let last = -1;
      let done = 0;
      const ordered = [...this.newestFirst].reverse();
      for (const [order, item] of ordered.entries()) {
        if (!selected.has(item.id)) continue;
        const identity: VersionIdentity = {
          label: `${this.target.name} · ${item.id}`, source: "cloud", number: order + 1,
          versionId: `cloud:${document}:${item.id}`, sourceModified: item.modified,
          gapBefore: last >= 0 && (order > last + 1 || item.cachedOnly === true || ordered[last].cachedOnly === true),
        };
        last = order;
        try {
          signal.throwIfAborted();
          versions.push(await this.importOne(dir, item, order, identity, signal));
        } catch (error) {
          const reason = signal.aborted ? "cancelled" : error instanceof CloudVersionError ? error.reason : "invalid";
          failures.push({ id: item.id, reason });
          versions.push({ ...identity, unavailable: true, threads: [], commentsUnavailable: true });
          if (reason === "sign-in") controller.abort();
        }
        progress(++done, selected.size);
      }
      if (controller.signal.aborted) return { versions, failures };
      if (!versions.some(version => !version.unavailable)) return { versions, failures };
      await this.verifyAccount();
      const bytes = this.target.bytes;
      const digest = hash(bytes);
      const snapshot = path.join(dir, `${digest}.docx`);
      const stored = await fs.readFile(snapshot).catch(() => undefined);
      if (!stored || hash(stored) !== digest) await replace(snapshot, bytes);
      const current = await readCloudSavedVersion(bytes, {
        label: this.target.name, source: "cloud", number: ordered.length + 1, sourceModified: this.target.mtime,
        versionId: `cloud:${document}:current:${digest}`,
        // The snapshot was opened independently of the list and has no Graph version ID.
        // Keep it searchable, but do not infer adjacency from timestamps or matching bytes.
        gapBefore: true,
      }, async () => { await this.verifyAccount(); return fs.readFile(snapshot); });
      return { versions: [...versions, current], failures };
    } finally {
      this.controller = undefined;
    }
  }

  private async importOne(dir: string, item: CloudVersionItem, order: number, identity: VersionIdentity, signal: AbortSignal) {
    const manifest = path.join(dir, `${hash(item.id)}.json`);
    const previous = await readManifest(manifest);
    const content = (digest: string) => path.join(dir, `${digest}.docx`);
    let bytes: Uint8Array | undefined = previous ? await fs.readFile(content(previous.sha256)).catch(() => undefined) : undefined;
    if (!bytes || hash(bytes) !== previous?.sha256) {
      if (item.cachedOnly) throw new CloudVersionError("invalid");
      const downloadSignal = AbortSignal.any([signal, AbortSignal.timeout(120_000)]);
      const received = await this.client.getVersionBytes(this.target, item, this.target.account, downloadSignal);
      const digest = hash(received);
      if (previous && digest !== previous.sha256) throw new CloudVersionError("conflict");
      try { await parseDocxLayout(received); } catch { throw new CloudVersionError("invalid"); }
      downloadSignal.throwIfAborted();
      await this.verifyAccount();
      await replace(content(digest), received);
      if (!previous) {
        await replace(manifest, JSON.stringify({ item, sha256: digest, fetchedAt: new Date().toISOString(), order }));
      }
      bytes = received;
    }
    const entry = await readManifest(manifest);
    if (!entry || !bytes) throw new CloudVersionError("invalid");
    return readCloudSavedVersion(bytes, { ...identity, fetchedAt: entry.fetchedAt }, async () => {
      await this.verifyAccount();
      return fs.readFile(content(entry.sha256));
    });
  }
}
