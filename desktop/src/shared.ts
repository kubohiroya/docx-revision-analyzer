/**
 * メインプロセスとレンダラの間でやり取りする型 (preload が window.app として公開する API)。
 *
 * 解析と描画はレンダラで、ライブラリのコア (docx-revision-analyzer/core。ファイルシステムを使わない) で行う。
 * メインプロセスは、ファイルの読み書き・ダイアログ・設定の保存・docx の設定の書き換えだけを受け持つ。
 * 文書や解析結果をネットワークへ送る経路は無い (main.ts で外部への通信を遮断している)。
 */

export type AppLang = "en" | "ja";

export interface AppSettings {
  /** 表示言語。undefined なら OS の言語 */
  lang?: AppLang;
  /** 判定ルールのファイル。undefined なら --bulk-chars 相当の既定ルール */
  rulesPath?: string;
  /** 拡張モジュールの有効/無効 (id → 有効か)。拡張の仕組みは #15 */
  extensions: Record<string, boolean>;
  /** 解析の設定 (前回の値を覚えておく) */
  analysis: AnalysisSettings;
}

export interface AnalysisSettings {
  /** 無編集期間がこれ (時間) を超えたら区間を分ける */
  gapThresholdHours: number;
  /** 一括挿入とみなす文字数 (ルールファイルが無いとき) */
  bulkChars: number;
  /** chart の時間の刻み ("auto" / "minute" / "hour" / "day" / 秒数) */
  bucket: string;
  /** chart を無編集期間で分けて描くか */
  chartSplit: boolean;
}

export const DEFAULT_SETTINGS: AppSettings = {
  extensions: {},
  analysis: { gapThresholdHours: 1, bulkChars: 150, bucket: "auto", chartSplit: true },
};

export interface OpenedFile {
  path: string;
  name: string;
  /** 最終更新日時 (ISO 8601) */
  mtime: string;
  bytes: Uint8Array;
}

export interface LoadedRules {
  path: string;
  /** YAML / JSON を解析した値 (検証はレンダラで parseRuleSet により行う) */
  value: unknown;
}

export interface AppApi {
  /** ファイルを開くダイアログ。キャンセルなら null */
  openDocxDialog(): Promise<OpenedFile | null>;
  /** パスのファイルを読む (ドロップされたファイルなど) */
  readDocx(path: string): Promise<OpenedFile>;
  /** ドロップされた File のパス */
  pathForFile(file: File): string;
  /** 変更履歴の作成者・日時が保存されるよう文書の設定を書き換える (元のファイルはバックアップ)。バックアップのパスを返す */
  preserveHistory(path: string): Promise<{ backupPath?: string }>;
  /** ルールファイルを選ぶダイアログ。キャンセルなら null */
  chooseRulesDialog(): Promise<LoadedRules | null>;
  /** 保存してあるパスのルールファイルを読む */
  readRules(path: string): Promise<LoadedRules>;
  /** 保存ダイアログを出してバイト列を書き出す。保存したパス (キャンセルなら null) を返す */
  saveFile(defaultName: string, data: Uint8Array | string, kind: "svg" | "png"): Promise<string | null>;
  getSettings(): Promise<AppSettings>;
  setSettings(settings: AppSettings): Promise<void>;
  /** OS の言語 */
  systemLang(): Promise<AppLang>;
  /** メインプロセスから「このファイルを開いて」と指示されたとき (Finder からのドロップ・スモークテスト) */
  onOpenPath(cb: (path: string) => void): void;
  /** アプリのバージョン */
  version(): Promise<string>;
}

declare global {
  interface Window {
    app: AppApi;
  }
}
