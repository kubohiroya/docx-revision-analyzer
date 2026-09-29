/**
 * ファイル名生成まわりのユーティリティ。
 * 主に「デスクトップからのドラッグ&ドロップ起動」モード (--drop) で、
 * 入力ファイルの最終更新日時をファイル名に埋め込むために使う。
 */
import * as path from "path";
import { t } from "./i18n";

function pad(n: number, width = 2): string {
  return String(n).padStart(width, "0");
}

/**
 * Date をファイル名に安全に埋め込める形式 (YYYYMMDD-HHMMSS、ローカル時刻) に整形する。
 * コロンはWindowsのファイル名で使えないため使用しない。
 */
export function formatTimestampForFilename(date: Date): string {
  const y = date.getFullYear();
  const mo = pad(date.getMonth() + 1);
  const d = pad(date.getDate());
  const h = pad(date.getHours());
  const mi = pad(date.getMinutes());
  const s = pad(date.getSeconds());
  return `${y}${mo}${d}-${h}${mi}${s}`;
}

/**
 * 入力ファイルパスと基準日時から、
 * "<同じディレクトリ>/<拡張子を除いたファイル名>-<YYYYMMDD-HHMMSS>.svg"
 * という出力パスを組み立てる。
 */
/** 図の見出しに添える日時 (ローカル時刻。英語は YYYY-MM-DD HH:MM、日本語は YYYY/MM/DD HH:MM) */
export function formatTimestampForTitle(date: Date): string {
  return t(
    "titleTime",
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    pad(date.getHours()),
    pad(date.getMinutes())
  );
}

/** 既定の図の見出し: "<接頭辞>: <ファイル名> (last modified / 最終更新 <日時>)" */
export function buildDefaultTitle(prefix: string, filePath: string, mtime: Date): string {
  return t("defaultTitle", prefix, path.basename(filePath), formatTimestampForTitle(mtime));
}

export function buildDropOutputPath(inputPath: string, mtime: Date, ext = ".svg"): string {
  const dir = path.dirname(inputPath);
  const base = path.basename(inputPath).replace(/\.[^.]+$/, "");
  const stamp = formatTimestampForFilename(mtime);
  return path.join(dir, `${base}-${stamp}${ext}`);
}
