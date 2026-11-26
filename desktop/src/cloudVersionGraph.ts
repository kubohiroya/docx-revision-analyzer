import { setTimeout as delay } from "node:timers/promises";

export const GRAPH_VERSION_ORIGIN = "https://graph.microsoft.com";
export const MAX_CLOUD_DOCX_BYTES = 200 * 1024 * 1024;

export interface CloudVersionItem {
  id: string;
  modified?: string;
  size?: number;
  cachedOnly?: boolean;
}

interface Target { driveId: string; itemId: string }

export class CloudVersionError extends Error {
  constructor(readonly reason: "sign-in" | "forbidden" | "missing" | "network" | "too-large" | "invalid" | "conflict" | "cancelled") {
    super(reason);
  }
}

export function cloudVersionPath(target: Target): string {
  return `/v1.0/drives/${encodeURIComponent(target.driveId)}/items/${encodeURIComponent(target.itemId)}/versions`;
}

function graphUrl(value: string, path: string): URL {
  const url = new URL(value, GRAPH_VERSION_ORIGIN);
  if (url.origin !== GRAPH_VERSION_ORIGIN || url.pathname !== path || url.username || url.password || url.hash) {
    throw new CloudVersionError("invalid");
  }
  return url;
}

async function request(url: URL, init: RequestInit, fetcher: typeof fetch): Promise<Response> {
  for (let attempt = 0; ; attempt++) {
    const response = await fetcher(url, { ...init, redirect: "manual" });
    if (response.status !== 429 || attempt === 2) return response;
    await response.body?.cancel();
    const retry = response.headers.get("retry-after");
    const seconds = retry && /^\d+$/.test(retry) ? Number(retry) : NaN;
    const ms = Number.isFinite(seconds) ? seconds * 1_000 : retry ? Date.parse(retry) - Date.now() : NaN;
    if (ms > 30_000) throw new CloudVersionError("network");
    await delay(Number.isFinite(ms) ? Math.max(0, ms) : 1_000 * (attempt + 1), undefined, { signal: init.signal ?? undefined });
  }
}

function failure(status: number): CloudVersionError {
  if (status === 401) return new CloudVersionError("sign-in");
  if (status === 403) return new CloudVersionError("forbidden");
  if (status === 404 || status === 410) return new CloudVersionError("missing");
  return new CloudVersionError("network");
}

export async function listCloudVersionItems(
  target: Target, token: string, signal: AbortSignal, fetcher: typeof fetch = fetch,
): Promise<CloudVersionItem[]> {
  const path = cloudVersionPath(target);
  let next: string | undefined = path;
  const visited = new Set<string>();
  const result = new Map<string, CloudVersionItem>();
  while (next) {
    const url = graphUrl(next, path);
    if (visited.has(url.href) || visited.size >= 100) throw new CloudVersionError("invalid");
    visited.add(url.href);
    const response = await request(url, { headers: { Authorization: `Bearer ${token}` }, signal }, fetcher);
    if (!response.ok) throw failure(response.status);
    const page = await response.json() as { value?: unknown; "@odata.nextLink"?: unknown };
    if (!Array.isArray(page.value)) throw new CloudVersionError("invalid");
    for (const value of page.value) {
      if (typeof value?.id !== "string" || !value.id || value.id === "." || value.id === "..") throw new CloudVersionError("invalid");
      if (!result.has(value.id)) result.set(value.id, {
        id: value.id,
        modified: typeof value.lastModifiedDateTime === "string" ? value.lastModifiedDateTime : undefined,
        size: typeof value.size === "number" && value.size >= 0 ? value.size : undefined,
      });
    }
    next = typeof page["@odata.nextLink"] === "string" ? page["@odata.nextLink"] : undefined;
  }
  return [...result.values()];
}

export async function getCloudVersionBytes(
  target: Target, item: CloudVersionItem, token: string, signal: AbortSignal, fetcher: typeof fetch = fetch,
): Promise<Uint8Array> {
  if ((item.size ?? 0) > MAX_CLOUD_DOCX_BYTES) throw new CloudVersionError("too-large");
  const path = `${cloudVersionPath(target)}/${encodeURIComponent(item.id)}/content`;
  let response = await request(graphUrl(path, path), { headers: { Authorization: `Bearer ${token}` }, signal }, fetcher);
  if (response.status === 302) {
    const location = response.headers.get("location");
    if (!location) throw new CloudVersionError("invalid");
    let download: URL;
    try { download = new URL(location); } catch { throw new CloudVersionError("invalid"); }
    if (download.protocol !== "https:" || download.username || download.password) throw new CloudVersionError("invalid");
    await response.body?.cancel();
    response = await request(download, { signal }, fetcher);
  }
  if (!response.ok) throw failure(response.status);
  if (!response.body) throw new CloudVersionError("invalid");
  if (Number(response.headers.get("content-length")) > MAX_CLOUD_DOCX_BYTES) {
    await response.body.cancel();
    throw new CloudVersionError("too-large");
  }
  const reader = response.body.getReader();
  const pieces: Uint8Array[] = [];
  let size = 0;
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > MAX_CLOUD_DOCX_BYTES) throw new CloudVersionError("too-large");
      pieces.push(value);
    }
  } finally {
    await reader.cancel().catch(() => undefined);
    reader.releaseLock();
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const piece of pieces) {
    bytes.set(piece, offset);
    offset += piece.byteLength;
  }
  return bytes;
}
