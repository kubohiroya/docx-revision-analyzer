/**
 * 拡張モジュール API v1 の型 (拡張の作者向け)。
 *
 * 拡張は1つのディレクトリで、manifest.json と JS のエントリ (ES モジュール) を持つ。
 * エントリは activate(host) を export する:
 *
 *   export function activate(host) {
 *     host.registerCategory({ id: "example.long", color: "#7C3AED", label: { en: "Long", ja: "長い" }, priority: 20 });
 *     host.onAnalysisComplete((r) => host.ui.showPanel({ id: "summary", title: "Summary", blocks: [...] }));
 *   }
 *
 * 拡張は、Node.js の権限もネットワークも持たない隔離されたプロセス (sandbox のレンダラ) で動き、
 * ここにある HostApi を通してだけアプリとやり取りする。詳しくは desktop/EXTENSIONS.md。
 */

/** このアプリが対応する API のバージョン */
export const SUPPORTED_API_VERSIONS = [1] as const;

export type LocalizedText = string | { ja?: string; en?: string };

export interface ExtensionManifest {
  /** 逆ドメイン形式の id (例: com.example.sample)。英小文字・数字・. - _ */
  id: string;
  name: LocalizedText;
  version: string;
  /** 対応する API のバージョン (いまは 1) */
  apiVersion: number;
  /** エントリの JS (ES モジュール)。ディレクトリからの相対パス */
  main: string;
  description?: LocalizedText;
  permissions?: ExtensionPermissions;
}

export interface ExtensionPermissions {
  /** 使う UI ("figure" は図の部分への注釈 = マウスオーバーの説明とポップアップ) */
  ui?: ("dialog" | "panel" | "form" | "figure")[];
  /** 図の注釈のリンク (クリックで既定のブラウザで開く) に使ってよいオリジン (https のみ) */
  links?: string[];
  /** net.post で送ってよい宛先のオリジン (https のみ) */
  network?: string[];
  /** 拡張ごとのローカル保存 */
  storage?: boolean;
  /** 解析結果に文書の本文 (挿入・削除された文章) を含めるか。無ければ本文は空文字列になる */
  documentText?: boolean;
}

// ---------------------------------------------------------------------------
// カテゴリと分類器 (#10, #11)
// ---------------------------------------------------------------------------

export interface CategorySpec {
  /** 拡張の id を接頭辞にすること (例: com.example.sample.long) */
  id: string;
  /** "#rrggbb" */
  color: string;
  label: LocalizedText;
  /** 既定のカテゴリは 0、判定ルールのレベルは 10 以上 */
  priority: number;
  pattern?: "hatch" | "cross" | "dots";
}

/** 分類器が返すハイライト (JSON にできる形。日時は ISO 8601 文字列) */
export interface HighlightSpec {
  categoryId: string;
  /** 対象の挿入の w:id */
  eventIds: string[];
  start: string;
  end: string;
  features?: Record<string, number | boolean>;
  reason: LocalizedText;
}

export interface ClassifierSpec {
  /** 拡張の id を接頭辞にすること */
  id: string;
  version: string;
  classify(ctx: ClassifierContext): HighlightSpec[] | Promise<HighlightSpec[]>;
}

/** 文書上の位置 (revisionPositions.ts と同じ) */
export interface Position {
  paraIndex: number;
  paraModelIndex: number;
  offsetInPara: number;
  docOffset: number;
  paraMissing: boolean;
}

export interface PositionedEvent {
  type: "ins" | "del";
  id: string;
  /** ISO 8601。日時が無ければ null */
  date: string | null;
  author: string | null;
  move: boolean;
  moveName: string | null;
  chars: number;
  /** 本文 (documentText の権限が無ければ空文字列) */
  text: string;
  includesParaMark: boolean;
  paraModelIndices: number[];
  final: { start: Position; end: Position };
  finalChars: number;
  atEdit: { start: Position; end: Position } | null;
}

export interface WindowOptions {
  seconds: number;
  chars?: number;
  paras?: number;
  byAuthor?: boolean;
}

export interface InsertionWindow {
  index: number;
  start: string;
  end: string;
  authors: string[];
  eventIds: string[];
  range: { start: Position; end: Position };
  features: Record<string, number | boolean>;
}

export interface ClassifierContext {
  positioned: PositionedEvent[];
  sessions: { start: string; end: string }[];
  /** 挿入の窓と特徴量 (#8) を求める (拡張の中で計算する) */
  windowsFor(options: WindowOptions): { timeResolutionSec: number; windows: InsertionWindow[] };
}

// ---------------------------------------------------------------------------
// 解析結果
// ---------------------------------------------------------------------------

export interface AnalysisResult {
  file: { name: string; mtime: string };
  /** 変更履歴の集計 (DocxRevisionData から events を除いたもの) */
  revisions: {
    eventCount: number;
    finalCharCount: number;
    baselineCharCount: number;
    totalInserted: number;
    totalDeleted: number;
    undatedRevisionCount: number;
    settings: { settingsPartFound: boolean; trackRevisions: boolean; removePersonalInformation: boolean; removeDateAndTime: boolean };
  };
  positioned: PositionedEvent[];
  sessions: { start: string; end: string; insChars: number; delChars: number; bulkInsChars: number; movedChars: number }[];
  /** 判定ルールの窓と、窓ごとのレベル */
  ruleSet: string;
  windows: (InsertionWindow & { level: string | null })[];
  /** すべての分類器のハイライト */
  highlights: (HighlightSpec & { classifierId: string })[];
  /** documentText の権限が無ければ空文字列 */
  finalText: string;
  /** 整合性の簡易チェック (integrity.ts。判定ではなく情報) */
  integrity: {
    items: { id: string; observed: boolean | null; details: Record<string, string | number | boolean | null>; message: LocalizedText }[];
    note: LocalizedText;
  };
}

// ---------------------------------------------------------------------------
// UI (宣言的な JSON。アプリが描画する)
// ---------------------------------------------------------------------------

export interface DialogSpec {
  title: LocalizedText;
  message: LocalizedText;
  /** 省略時は OK だけ */
  buttons?: { id: string; label: LocalizedText; primary?: boolean }[];
}

export interface DialogResult {
  /** 押されたボタンの id。閉じられた場合は null */
  buttonId: string | null;
}

export type PanelBlock =
  | { type: "heading"; text: LocalizedText }
  | { type: "text"; text: LocalizedText }
  | { type: "list"; items: LocalizedText[] }
  | { type: "keyValue"; rows: { key: LocalizedText; value: string | number }[] }
  | { type: "table"; columns: LocalizedText[]; rows: (string | number)[][] };

export interface PanelSpec {
  /** 同じ id のパネルは置き換える */
  id: string;
  title: LocalizedText;
  blocks: PanelBlock[];
}

export type FormField =
  | { id: string; type: "text" | "textarea"; label: LocalizedText; required?: boolean; default?: string; maxLength?: number }
  | { id: string; type: "number"; label: LocalizedText; required?: boolean; default?: number; min?: number; max?: number }
  | { id: string; type: "checkbox"; label: LocalizedText; default?: boolean }
  | { id: string; type: "select"; label: LocalizedText; options: { value: string; label: LocalizedText }[]; default?: string };

export interface FormSpec {
  title: LocalizedText;
  description?: LocalizedText;
  fields: FormField[];
  submitLabel?: LocalizedText;
}

/** 送信された値 (キャンセルなら openForm は null を返す) */
export type FormResult = Record<string, string | number | boolean>;

// ---------------------------------------------------------------------------
// 図の部分への注釈
// ---------------------------------------------------------------------------

/**
 * 図の部分 (figureTargets.ts と同じ)。key を注釈の target に使う。
 *  - chart の棒: chart:bar:<カテゴリ id>:<バケットの開始時刻>
 *  - flow の段落: flow:para:<列>:<段落>、帯: flow:band:<区間>:<単位>、移動の帯: flow:move:...、キャプション: flow:caption:<区間>
 */
export type FigureTarget = {
  /** 段落の冒頭 (約 40 文字。documentText の権限がある拡張だけに渡す)。段落・帯だけ */
  excerpt?: string;
  /** 段落を含むセクションの名前 (それ以前で最後の見出しの文字、またはブックマークの名前)。段落・帯だけ */
  section?: string;
  /**
   * 元の文書の該当箇所へのリンク (OneDrive / SharePoint の文書のときだけ)。見出しのセクションなら Word for the web の
   * 見出しリンク (nav=)、ブックマークなら URL#ブックマーク名、どちらも無ければ文書の URL (冒頭)。
   * href にそのまま使えば links の権限は要らない
   */
  docLink?: string;
} & (
  | { key: string; figure: "chart"; kind: "bar"; categoryId: string; start: string; end: string; chars: number }
  | { key: string; figure: "flow"; kind: "paragraph"; column: number; paraIndex: number; session: number | null }
  | { key: string; figure: "flow"; kind: "band"; session: number; unitKey: string; paraIndexes: number[]; change: string }
  | { key: string; figure: "flow"; kind: "move"; session: number; fromKey: string; toKey: string; chars: number }
  | { key: string; figure: "flow"; kind: "caption"; session: number; start: string; end: string }
);

export interface FigureContext {
  figure: "chart" | "flow";
  targets: FigureTarget[];
  analysis: AnalysisResult;
}

export interface FigureAnnotationSpec {
  /** 部分のキー (FigureTarget.key) */
  target: string;
  /** マウスオーバーで出る短い説明 (書き出した SVG では <title> になる) */
  tooltip?: LocalizedText;
  /** マウスオーバーで出るポップアップの中身 (アプリの中だけ。パネルと同じブロック) */
  popup?: { title?: LocalizedText; blocks: PanelBlock[] };
  /** クリックで開くリンク (https。permissions.links で宣言したオリジンだけ) */
  href?: string;
}

export interface FigureAnnotatorSpec {
  /** 拡張の id を接頭辞にすること */
  id: string;
  version: string;
  annotate(ctx: FigureContext): FigureAnnotationSpec[] | Promise<FigureAnnotationSpec[]>;
}

// ---------------------------------------------------------------------------
// ホスト API
// ---------------------------------------------------------------------------

export interface KeyValueStore {
  get<T = unknown>(key: string): Promise<T | undefined>;
  set(key: string, value: unknown): Promise<void>;
  delete(key: string): Promise<void>;
  keys(): Promise<string[]>;
}

export interface NetResponse {
  ok: boolean;
  status: number;
  body: string;
  /** オフライン等で送れず、キューに入れた (つながったらアプリが送る)。このとき ok は false、status は 0 */
  queued?: boolean;
}

export interface PostOptions {
  /** true なら、利用者の設定にかかわらず送信前に確認を求める */
  confirm?: boolean;
}

export interface HostApi {
  registerCategory(c: CategorySpec): void;
  registerClassifier(c: ClassifierSpec): void;
  /** 図の部分に、マウスオーバーの説明・ポップアップ・クリックで開くリンクを付ける (ui: "figure" の権限が必要) */
  registerFigureAnnotator(a: FigureAnnotatorSpec): void;
  onAnalysisComplete(cb: (r: AnalysisResult) => void | Promise<void>): void;
  ui: {
    showDialog(d: DialogSpec): Promise<DialogResult>;
    showPanel(p: PanelSpec): void;
    openForm(f: FormSpec): Promise<FormResult | null>;
  };
  /** 拡張ごとのローカル保存 (storage の権限が必要) */
  storage: KeyValueStore;
  /**
   * manifest で許可したオリジンにだけ JSON を POST する (network の権限が必要)。
   * アプリが仲介し、送信履歴に記録する。利用者が送信前の確認を有効にしている場合 (既定) や
   * options.confirm のときは、宛先と本文を見せて確認する (拒否されたら例外)。
   */
  net: { post(url: string, body: unknown, options?: PostOptions): Promise<NetResponse> };
  app: { version: string; locale: "ja" | "en"; apiVersion: number };
}

export interface Extension {
  activate(host: HostApi): void | Promise<void>;
}
