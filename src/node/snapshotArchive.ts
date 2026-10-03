/**
 * snapshotArchive.ts (Node.js 専用)
 *
 * 区切りごとに提出された文書 (スナップショット) を、通し番号を付けて保管する。
 *
 *   <保管先>/<キー>/snapshots.json        … 保管した回の一覧 (提出元・保管日時・SHA-256)
 *   <保管先>/<キー>/<名前>-s01.docx        … 第1回の提出物 (以降 -s02, -s03 ...)
 *
 * キーは文書ごとのフォルダ名で、入力したフォルダからの相対パス (拡張子を除く) から作る。
 * 毎回同じフォルダ構成で提出物を渡せば、同じ文書は同じキーになる。
 */

import * as fs from "fs";
import * as path from "path";
import { createHash } from "crypto";

export const MANIFEST_NAME = "snapshots.json";

export interface SnapshotEntry {
  /** 通し番号 (1 から) */
  n: number;
  /** 保管したファイル名 (キーのフォルダの中) */
  file: string;
  /** 提出元のパス */
  source: string;
  /** 提出元のファイルの最終更新日時 (ISO 8601) */
  sourceModified: string;
  /** 保管した日時 (ISO 8601) */
  archivedAt: string;
  sha256: string;
}

export interface SnapshotManifest {
  version: 1;
  key: string;
  snapshots: SnapshotEntry[];
}

/**
 * 入力の相対パス (例: "田中/卒論.docx") から文書のキーを作る。
 * keyDepth を指定すると、先頭からその数のフォルダだけをキーにする (提出のたびにファイル名や途中のフォルダ名が
 * 変わる場合、例えば "<学生>/<課題>/<ファイル>" の形なら 1)。フォルダがそれより浅ければ、拡張子を除いたパス全体
 */
export function snapshotKey(relativePath: string, keyDepth?: number): string {
  const parts = relativePath.split(/[/\\]/).filter((p) => p && p !== ".");
  const dirs = parts.slice(0, -1);
  if (keyDepth !== undefined && keyDepth > 0 && dirs.length >= keyDepth) return dirs.slice(0, keyDepth).join("/");
  return [...dirs, parts[parts.length - 1].replace(/\.docx$/i, "")].join("/");
}

/** キーのフォルダ (保管先の下) */
export function keyDir(archive: string, key: string): string {
  return path.join(archive, ...key.split("/"));
}

/** 保管するファイル名の元 (キーの最後の部分) */
export function keyBaseName(key: string): string {
  return key.split("/").pop() || "document";
}

export function readManifest(dir: string): SnapshotManifest | undefined {
  const file = path.join(dir, MANIFEST_NAME);
  if (!fs.existsSync(file)) return undefined;
  const m = JSON.parse(fs.readFileSync(file, "utf-8")) as SnapshotManifest;
  if (!Array.isArray(m.snapshots)) throw new Error(`${file}: invalid manifest`);
  return m;
}

function writeManifest(dir: string, m: SnapshotManifest): void {
  fs.writeFileSync(path.join(dir, MANIFEST_NAME), JSON.stringify(m, null, 2) + "\n", "utf-8");
}

export function snapshotFileName(key: string, n: number): string {
  return `${keyBaseName(key)}-s${String(n).padStart(2, "0")}.docx`;
}

export interface AddSnapshotResult {
  entry: SnapshotEntry;
  /** 新しく保管したか (直前の回と同じ内容なら false) */
  added: boolean;
  dir: string;
}

/** 文書を次の番号で保管する。直前の回と同じ内容 (SHA-256 が同じ) なら保管しない */
export async function addSnapshot(archive: string, key: string, source: string, now = new Date()): Promise<AddSnapshotResult> {
  const dir = keyDir(archive, key);
  await fs.promises.mkdir(dir, { recursive: true });
  const bytes = await fs.promises.readFile(source);
  const sha256 = createHash("sha256").update(bytes).digest("hex");
  const m = readManifest(dir) ?? { version: 1, key, snapshots: [] };
  const last = m.snapshots[m.snapshots.length - 1];
  if (last && last.sha256 === sha256) return { entry: last, added: false, dir };

  const n = (last?.n ?? 0) + 1;
  const file = snapshotFileName(key, n);
  // 既にあるファイルは上書きしない (一覧と食い違っている場合は、手で確かめてもらう)
  await fs.promises.writeFile(path.join(dir, file), bytes, { flag: "wx" });
  const entry: SnapshotEntry = {
    n,
    file,
    source: path.resolve(source),
    sourceModified: fs.statSync(source).mtime.toISOString(),
    archivedAt: now.toISOString(),
    sha256,
  };
  m.snapshots.push(entry);
  writeManifest(dir, m);
  return { entry, added: true, dir };
}

/** 保管先の下の、スナップショットを保管しているキーの一覧 */
export function listSnapshotKeys(archive: string): string[] {
  const keys: string[] = [];
  const walk = (dir: string, rel: string[]) => {
    let entries: fs.Dirent[];
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    if (entries.some((e) => e.isFile() && e.name === MANIFEST_NAME)) keys.push(rel.join("/"));
    for (const e of entries.sort((a, b) => a.name.localeCompare(b.name))) {
      if (e.isDirectory() && !e.name.startsWith(".")) walk(path.join(dir, e.name), [...rel, e.name]);
    }
  };
  walk(archive, []);
  return keys.filter((k) => k !== "");
}
