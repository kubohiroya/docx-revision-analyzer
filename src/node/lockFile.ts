/**
 * lockFile.ts (Node.js 専用)
 *
 * テンプレートへの「変更履歴のロック」(trackLock.ts) を、node:crypto のハッシュ関数と乱数で行い、
 * ファイルに適用する。元のファイルはバックアップとして残す。
 */

import * as fs from "fs";
import { createHash, randomBytes } from "crypto";
import { LockOptions, LockResult, lockTrackChangesInDocx, verifyLockPassword } from "../lib/trackLock";
import type { DocxInput } from "../lib/input";
import { checkNotOpenElsewhere, replaceWithBackup } from "./historyFile";

/** node:crypto の SHA-512 (Web Crypto より速い。10 万回のハッシュを同期的に行う) */
export const nodeSha512 = (data: Uint8Array): Uint8Array => createHash("sha512").update(data).digest();

/** docx のバイト列に、node:crypto でロックをかける */
export function lockDocxBytes(input: DocxInput, password: string | undefined): Promise<LockResult> {
  const opts: LockOptions = { password, digest: nodeSha512, salt: new Uint8Array(randomBytes(16)) };
  return lockTrackChangesInDocx(input, opts);
}

/** ロックのパスワードが合っているか (node:crypto で確かめる) */
export function verifyLockPasswordNode(input: DocxInput, password: string): Promise<boolean> {
  return verifyLockPassword(input, password, nodeSha512);
}

export interface LockFileResult extends LockResult {
  /** 書き換える前のファイルのバックアップ先 */
  backupPath: string;
}

/**
 * テンプレートのファイルに変更履歴のロックをかけて上書きする。
 * 元のファイルは "<名前>.backup-<YYYYMMDD-HHMMSS>.docx" として同じフォルダに残す。
 * ファイルが他で開かれている場合はエラーを投げる。
 */
export async function lockTemplateFile(filePath: string, password: string | undefined): Promise<LockFileResult> {
  const openReason = checkNotOpenElsewhere(filePath);
  if (openReason) throw new Error(openReason);
  const r = await lockDocxBytes(await fs.promises.readFile(filePath), password);
  const backupPath = await replaceWithBackup(filePath, r.output);
  return { ...r, backupPath };
}
