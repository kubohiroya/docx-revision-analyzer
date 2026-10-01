/**
 * historyFile.ts (Node.js 専用)
 *
 * 変更履歴の設定の書き換え (historySettings.ts の preserveHistoryInDocx) をファイルに適用する。
 * Word 等で開かれていないかの確認、バックアップ、一時ファイルを使った置き換えを行う。
 */

import * as fs from "fs";
import * as path from "path";
import { spawnSync } from "child_process";
import {
  DocxRevisionSettings,
  needsHistoryFix,
  preserveHistoryInDocx,
  readHistorySettings,
} from "../lib/historySettings";
import { formatTimestampForFilename } from "../lib/filenames";
import { t } from "../lib/i18n";

// ---------------------------------------------------------------------------
// ファイルが他で開かれていないかの確認
// ---------------------------------------------------------------------------

/**
 * Word が文書を開いている間に同じフォルダへ作る所有者ファイル (~$xxxx.docx) を探す。
 * Word はファイル名が長い場合、先頭の1〜2文字を "~$" で置き換えた名前にする。
 */
function findWordLockFile(filePath: string): string | undefined {
  const dir = path.dirname(filePath);
  const name = path.basename(filePath).normalize("NFC");
  let entries: string[];
  try {
    entries = fs.readdirSync(dir);
  } catch {
    return undefined;
  }
  for (const entry of entries) {
    const e = entry.normalize("NFC");
    if (!e.startsWith("~$")) continue;
    const rest = e.slice(2);
    if (rest.length > 0 && rest.length >= name.length - 2 && name.endsWith(rest)) {
      return path.join(dir, entry);
    }
  }
  return undefined;
}

/** 他のプロセスがファイルを開いているか (OS ごとのベストエフォート) */
function isOpenByAnotherProcess(filePath: string): boolean {
  if (process.platform === "win32") {
    // Windows では Word が開いているファイルは共有違反で書き込みオープンできない
    try {
      fs.closeSync(fs.openSync(filePath, "r+"));
      return false;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      return code === "EBUSY" || code === "EPERM" || code === "EACCES";
    }
  }
  // macOS / Linux: lsof が使えれば、ファイルを開いているプロセスの有無を調べる
  const r = spawnSync("lsof", ["-t", "--", filePath], { encoding: "utf-8" });
  if (r.error) return false;
  return r.stdout.trim().length > 0;
}

/**
 * ファイルが Word 等で開かれていないことを確認する。
 * 開かれている可能性がある場合は、その理由を説明する文字列を返す。
 */
export function checkNotOpenElsewhere(filePath: string): string | undefined {
  const lock = findWordLockFile(filePath);
  if (lock) {
    return t("openInWord", path.basename(lock));
  }
  if (isOpenByAnotherProcess(filePath)) {
    return t("openElsewhere");
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// ファイルへの適用
// ---------------------------------------------------------------------------

export interface EnableHistoryResult {
  /** 設定を書き換えたか (もともと問題が無ければ false) */
  changed: boolean;
  /** 書き換え前のファイルのバックアップ先 */
  backupPath?: string;
  before: DocxRevisionSettings;
}

/**
 * 文書の設定を書き換え、変更履歴の作成者・日時が保存されるようにする。
 * 元のファイルは "<名前>.backup-<YYYYMMDD-HHMMSS>.docx" として同じフォルダに残す。
 * ファイルが他で開かれている場合はエラーを投げる。
 */
export async function enableHistoryPreservation(filePath: string): Promise<EnableHistoryResult> {
  const buf = await fs.promises.readFile(filePath);
  const before = await readHistorySettings(buf);
  if (!before.settingsPartFound || !needsHistoryFix(before)) return { changed: false, before };

  const openReason = checkNotOpenElsewhere(filePath);
  if (openReason) throw new Error(openReason);

  const r = await preserveHistoryInDocx(buf);
  if (!r.output) return { changed: false, before };
  const out = r.output;

  const dir = path.dirname(filePath);
  const base = path.basename(filePath).replace(/\.docx$/i, "");
  const backupPath = path.join(dir, `${base}.backup-${formatTimestampForFilename(new Date())}.docx`);
  await fs.promises.copyFile(filePath, backupPath, fs.constants.COPYFILE_EXCL);

  // 途中で失敗しても元のファイルが壊れないよう、一時ファイルに書いてから置き換える
  const tmpPath = path.join(dir, `.${base}.tmp-${process.pid}.docx`);
  try {
    await fs.promises.writeFile(tmpPath, out);
    await fs.promises.rename(tmpPath, filePath);
  } catch (err) {
    await fs.promises.rm(tmpPath, { force: true });
    throw err;
  }
  return { changed: true, backupPath, before };
}
