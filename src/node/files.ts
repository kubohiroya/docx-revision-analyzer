/**
 * files.ts (Node.js 専用)
 *
 * ファイルを読み書きする薄いラッパー。解析そのものはバイト列を受け取るコア (src/lib) で行う。
 */

import * as fs from "fs";
import * as path from "path";
import { DocxRevisionData, ExtractOptions, extractRevisions } from "../lib/docxRevisions";
import { DocxLayoutModel, parseDocxLayout } from "../lib/docxLayout";
import { formatTimestampForFilename } from "../lib/filenames";

export async function extractRevisionsFromFile(
  filePath: string,
  opts: ExtractOptions = {}
): Promise<DocxRevisionData> {
  return extractRevisions(await fs.promises.readFile(filePath), opts);
}

export async function parseDocxLayoutFromFile(filePath: string): Promise<DocxLayoutModel> {
  return parseDocxLayout(await fs.promises.readFile(filePath));
}

/**
 * 入力ファイルパスと基準日時から、
 * "<同じディレクトリ>/<拡張子を除いたファイル名>-<YYYYMMDD-HHMMSS>.svg"
 * という出力パスを組み立てる。
 */
export function buildDropOutputPath(inputPath: string, mtime: Date, ext = ".svg"): string {
  const dir = path.dirname(inputPath);
  const base = path.basename(inputPath).replace(/\.[^.]+$/, "");
  const stamp = formatTimestampForFilename(mtime);
  return path.join(dir, `${base}-${stamp}${ext}`);
}
