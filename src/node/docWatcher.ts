/**
 * docWatcher.ts (Node.js 専用)
 *
 * 1つの文書 (.docx) を監視し、保存されるたびに次のことを行う (デスクトップアプリの「監視」)。
 *  - 保存ごとの版を、文書ごとの記録のフォルダに通し番号で保管する (docx-revision-snapshot のアーカイブと同じ形式)
 *  - Word のコメントをすべて読み、出来事の列 (comments.jsonl、追記だけ) に積む
 *  - 新しく見たコメントへの「保全済み」の返信を待ち行列にためる
 * 文書が閉じられたこと (Word の所有者ファイル ~$… が無く、Windows では Word がファイルをつかんでいない) を
 * 検知したら、待ち行列の返信を文書に書き込む。監視中 (文書が開いている間) は文書に書き込まない。
 *
 * 読み取り元 (WatchSource) は、ローカルのファイル (fileSource: フォルダの変更の通知と、更新日時の確認) のほか、
 * 差し替えられる (デスクトップアプリは OneDrive / SharePoint の文書を Graph への問い合わせで監視する)。
 * 書き込めない読み取り元 (canWrite が false) では、返信を待ち行列に入れない。
 *
 * 記録のフォルダ (<記録の保管先>/<文書の鍵>/):
 *   watch.json        対象のパス、文書の識別子 (w15:docId)、監視した期間、監視していない間の変化、ツールの書き込み
 *   snapshots.json    保存ごとの版の一覧 / <名前>-sNN.docx 保存ごとの版
 *   comments.jsonl    コメントの出来事の列 (1行1件の JSON)
 *   receipts.json     「保全済み」の返信の待ち行列と、書き込みの結果
 */

import * as fs from "fs";
import * as path from "path";
import { createHash } from "crypto";
import JSZip from "jszip";
import {
  addCommentReplies,
  CommentEvent,
  CommentState,
  diffComments,
  extractComments,
  replayCommentEvents,
} from "../lib/docxComments";
import { addSnapshotBytes, readManifest, SnapshotEntry } from "./snapshotArchive";

export interface WatchInfo {
  version: 1;
  /** 監視している文書のパス (OneDrive / SharePoint の文書は URL) */
  path: string;
  /** これまでに監視したパス (移動・改名を追ったもの) */
  paths: string[];
  /** 文書の識別子 (word/settings.xml の w15:docId) */
  docId?: string;
  /** 監視した期間。end が無いのは監視中 (またはアプリが異常終了した) */
  periods: { start: string; end?: string }[];
  /** 監視していない間に文書が変わっていたことの記録 */
  gaps: { since?: string; detectedAt: string; sha256: string }[];
  /** 最後に読んだ版 */
  lastSha256?: string;
  lastSeenAt?: string;
  /** ツールが書き込んだ版 (新しい版として扱わない) */
  toolWrites: { sha256: string; at: string }[];
}

export interface ReceiptQueue {
  version: 1;
  /** 書き込みを待っている返信 (コメントの鍵と記録番号) */
  pending: { key: string; seq: number; queuedAt: string }[];
  /** 書き込んだ返信 */
  written: { key: string; seq: number; at: string; sha256: string }[];
  /** 書き込めなかった返信 (コメントが消えていた・本文に範囲の印が無い) */
  skipped: { key: string; seq: number; at: string; reason: string }[];
}

const INFO = "watch.json";
const EVENTS = "comments.jsonl";
const RECEIPTS = "receipts.json";

function sha256Of(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function readJson<T>(file: string): T | undefined {
  try {
    return JSON.parse(fs.readFileSync(file, "utf-8")) as T;
  } catch {
    return undefined;
  }
}

function writeJson(file: string, value: unknown): void {
  // 書きかけの状態を残さないよう、一時ファイルに書いてから置き換える
  const tmp = `${file}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value, null, 2) + "\n", "utf-8");
  fs.renameSync(tmp, file);
}

/** 記録のフォルダ名: 文書の名前 (拡張子なし) と、パス (または URL) のハッシュ値 */
export function watchKey(location: string, name = path.basename(location)): string {
  const base = name.replace(/\.docx$/i, "").replace(/[\\/:*?"<>|]/g, "_").slice(0, 40) || "document";
  const abs = /^https?:/i.test(location) ? location : path.resolve(location);
  return `${base}-${sha256Of(new TextEncoder().encode(abs)).slice(0, 10)}`;
}

/** 文書の識別子 (w15:docId)。無ければ undefined */
export async function docIdOf(bytes: Uint8Array): Promise<string | undefined> {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const settings = await zip.file("word/settings.xml")?.async("string");
    return settings?.match(/<w15:docId\b[^>]*w15:val="([^"]+)"/)?.[1];
  } catch {
    return undefined;
  }
}

/** 文書ごとの記録のフォルダ */
export class WatchArchive {
  readonly dir: string;
  readonly key: string;
  info: WatchInfo;

  private constructor(
    readonly root: string,
    key: string,
    info: WatchInfo
  ) {
    this.key = key;
    this.dir = path.join(root, key);
    this.info = info;
  }

  /**
   * 文書の記録のフォルダを開く (無ければ作る)。このパスの記録が無くても、同じ識別子 (w15:docId) で最後の版が
   * 同じ内容の記録があれば、移動・改名した文書とみなしてその記録を使う
   */
  static open(root: string, location: string, current?: { docId?: string; sha256?: string; name?: string }): WatchArchive {
    const abs = /^https?:/i.test(location) ? location : path.resolve(location);
    let key = watchKey(abs, current?.name);
    let info = readJson<WatchInfo>(path.join(root, key, INFO));
    if (!info && current?.docId && current.sha256) {
      for (const d of fs.existsSync(root) ? fs.readdirSync(root) : []) {
        const other = readJson<WatchInfo>(path.join(root, d, INFO));
        if (other?.docId === current.docId && other.lastSha256 === current.sha256) {
          key = d;
          info = other;
          break;
        }
      }
    }
    info ??= { version: 1, path: abs, paths: [], periods: [], gaps: [], toolWrites: [] };
    if (!info.paths.includes(abs)) info.paths.push(abs);
    info.path = abs;
    if (current?.docId) info.docId = current.docId;
    fs.mkdirSync(path.join(root, key), { recursive: true });
    const a = new WatchArchive(root, key, info);
    a.saveInfo();
    return a;
  }

  saveInfo(): void {
    writeJson(path.join(this.dir, INFO), this.info);
  }

  /** 保存ごとの版を保管する。直前の版と同じ内容なら保管しない */
  async addVersion(bytes: Uint8Array, mtime: Date, now: Date): Promise<SnapshotEntry | undefined> {
    const r = await addSnapshotBytes(this.root, this.key, bytes, { source: this.info.path, sourceModified: mtime, now });
    return r.added ? r.entry : undefined;
  }

  versions(): SnapshotEntry[] {
    return readManifest(this.dir)?.snapshots ?? [];
  }

  readEvents(): CommentEvent[] {
    const file = path.join(this.dir, EVENTS);
    if (!fs.existsSync(file)) return [];
    return fs
      .readFileSync(file, "utf-8")
      .split("\n")
      .filter((l) => l.trim())
      .flatMap((l) => {
        try {
          return [JSON.parse(l) as CommentEvent];
        } catch {
          // 書きかけの行 (異常終了など) は読み飛ばす
          return [];
        }
      });
  }

  /** 出来事に通し番号を付けて記録に追記する */
  appendEvents(events: Omit<CommentEvent, "seq">[]): CommentEvent[] {
    if (events.length === 0) return [];
    const existing = this.readEvents();
    let seq = existing.length ? existing[existing.length - 1].seq : 0;
    const out = events.map((e) => ({ seq: ++seq, ...e }) as CommentEvent);
    fs.appendFileSync(path.join(this.dir, EVENTS), out.map((e) => JSON.stringify(e)).join("\n") + "\n", "utf-8");
    return out;
  }

  commentState(): Map<string, CommentState> {
    return replayCommentEvents(this.readEvents());
  }

  receipts(): ReceiptQueue {
    return readJson<ReceiptQueue>(path.join(this.dir, RECEIPTS)) ?? { version: 1, pending: [], written: [], skipped: [] };
  }

  saveReceipts(q: ReceiptQueue): void {
    writeJson(path.join(this.dir, RECEIPTS), q);
  }
}

// ---------------------------------------------------------------------------
// Word で開いているかの判定
// ---------------------------------------------------------------------------

/**
 * Word が文書を開いているか。Word は開いている文書と同じフォルダに所有者ファイル (~$ で始まる名前。文書の名前が
 * 長いときは先頭の1〜2文字を ~$ に置き換えたもの) を作る。Windows では、Word が開いている文書は書き込み用に開けない
 */
export async function isOpenInWord(docPath: string): Promise<boolean> {
  const dir = path.dirname(docPath);
  const name = path.basename(docPath);
  let entries: string[] = [];
  try {
    entries = await fs.promises.readdir(dir);
  } catch {
    return false;
  }
  const owner = entries.some((e) => {
    if (!e.startsWith("~$")) return false;
    const rest = e.slice(2);
    return rest.length >= name.length - 2 && rest.length > 0 && name.endsWith(rest);
  });
  if (owner) return true;
  if (process.platform === "win32") {
    try {
      const h = await fs.promises.open(docPath, "r+");
      await h.close();
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === "EBUSY" || code === "EPERM" || code === "EACCES") return true;
    }
  }
  return false;
}

/** ZIP として完全に読める .docx か (保存の途中のファイルを読まないため) */
async function isCompleteDocx(bytes: Uint8Array): Promise<boolean> {
  try {
    const zip = await JSZip.loadAsync(bytes);
    return !!zip.file("word/document.xml");
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// 読み取り元
// ---------------------------------------------------------------------------

export interface WatchSource {
  /** 文書の場所 (ファイルのパス、または URL)。記録のフォルダを決めるのに使う */
  location: string;
  name: string;
  /** 文書を読む (保存の途中のファイルは読まない) */
  read(): Promise<{ bytes: Uint8Array; mtime: Date }>;
  /** 変化の知らせを受け取り始める。やめる関数を返す */
  subscribe(onChange: () => void, onError: (err: unknown) => void): () => void;
  /** 「保全済み」の返信を書き込めるか */
  canWrite: boolean;
  /** 文書が開かれているか (書き込む前に確かめる) */
  isOpen?(): Promise<boolean>;
  /** 文書を書き換える */
  write?(bytes: Uint8Array): Promise<void>;
}

/** ローカルのファイルの読み取り元 */
export function fileSource(
  docPath: string,
  opts: { debounceMs?: number; statPollMs?: number; isOpen?: (p: string) => Promise<boolean> } = {}
): WatchSource {
  const abs = path.resolve(docPath);
  const statKey = () => {
    try {
      const st = fs.statSync(abs);
      return `${st.mtimeMs}:${st.size}`;
    } catch {
      return "";
    }
  };
  let lastStat = statKey();
  return {
    location: abs,
    name: path.basename(abs),
    canWrite: true,
    async read() {
      let last: unknown;
      for (let i = 0; i < 6; i++) {
        try {
          const bytes = new Uint8Array(await fs.promises.readFile(abs));
          if (await isCompleteDocx(bytes)) {
            lastStat = statKey();
            return { bytes, mtime: fs.statSync(abs).mtime };
          }
          last = new Error(`${abs}: not a complete .docx`);
        } catch (err) {
          last = err;
        }
        // ZIP として完全に読めるまで、少し待って読み直す (Word が保存している途中など)
        await new Promise((r) => setTimeout(r, 300 * (i + 1)));
      }
      throw last;
    },
    subscribe(onChange, onError) {
      let debounce: NodeJS.Timeout | undefined;
      const schedule = () => {
        if (debounce) clearTimeout(debounce);
        debounce = setTimeout(onChange, opts.debounceMs ?? 1500);
      };
      let watcher: fs.FSWatcher | undefined;
      try {
        // Word は一時ファイルに書いてから名前を変えることがあるので、フォルダを監視して名前で絞る
        watcher = fs.watch(path.dirname(abs), (_event, filename) => {
          if (filename === null || filename.toString() === path.basename(abs)) schedule();
        });
        watcher.on("error", onError);
      } catch (err) {
        onError(err);
      }
      // ファイルの変更の通知が届かない場所 (ネットワークのドライブなど) のため、更新日時と大きさも確かめる
      const timer = setInterval(() => {
        const k = statKey();
        if (k !== lastStat) {
          lastStat = k;
          schedule();
        }
      }, opts.statPollMs ?? 5000);
      return () => {
        watcher?.close();
        clearInterval(timer);
        if (debounce) clearTimeout(debounce);
      };
    },
    isOpen: () => (opts.isOpen ?? isOpenInWord)(abs),
    async write(bytes) {
      // 書きかけのファイルを読まれないよう、同じフォルダの一時ファイルに書いてから置き換える
      const tmp = path.join(path.dirname(abs), `.~dra-${path.basename(abs)}.tmp`);
      await fs.promises.writeFile(tmp, bytes);
      await fs.promises.rename(tmp, abs);
      lastStat = statKey();
    },
  };
}

// ---------------------------------------------------------------------------
// 監視
// ---------------------------------------------------------------------------

export type WatchUpdateKind = "initial" | "saved" | "catchUp";

export interface WatchUpdate {
  kind: WatchUpdateKind;
  /** 文書の場所 (ファイルのパス、または URL) */
  path: string;
  bytes: Uint8Array;
  sha256: string;
  mtime: Date;
  /** 新しく保管した版 (同じ内容なら undefined) */
  version?: SnapshotEntry;
  /** この読み取りで記録した出来事 */
  events: CommentEvent[];
  /** すべてのコメントの状態 (記録番号の順) */
  comments: CommentState[];
  /** 監視していない間に文書が変わっていた */
  gap: boolean;
}

export interface ReceiptsWritten {
  path: string;
  written: { key: string; seq: number }[];
  skipped: { key: string; seq: number; reason: string }[];
  sha256: string;
}

export interface DocWatcherOptions {
  /** 監視するローカルの文書 (source を指定しないとき) */
  path?: string;
  /** 読み取り元 (OneDrive / SharePoint の文書など)。指定しなければ path のファイル */
  source?: WatchSource;
  /** 記録の保管先 (この下に文書ごとのフォルダを作る) */
  recordsRoot: string;
  /** 「保全済み」の返信の文面 */
  receiptText: (seq: number, at: Date) => string;
  onUpdate: (u: WatchUpdate) => void;
  onReceipts?: (r: ReceiptsWritten) => void;
  onError?: (err: unknown) => void;
  /** ローカルのファイル: 最後の変化からこれだけ待ってから読む (自動保存の連続を考えて) */
  debounceMs?: number;
  /** 閉じたかどうかを確かめる間隔 (返信の待ち行列があるときだけ) */
  closePollMs?: number;
  /** ローカルのファイル: 更新日時・大きさを確かめる間隔 (ファイルの変更の通知が届かない場所のため) */
  statPollMs?: number;
  /** ローカルのファイル: 開いているかの判定 (テスト用に差し替えられる) */
  isOpen?: (docPath: string) => Promise<boolean>;
  now?: () => Date;
}

export class DocWatcher {
  archive!: WatchArchive;
  readonly source: WatchSource;
  private unsubscribe?: () => void;
  private closeTimer?: NodeJS.Timeout;
  private busy: Promise<unknown> = Promise.resolve();
  private stopped = false;
  private readonly now: () => Date;

  constructor(private readonly o: DocWatcherOptions) {
    this.now = o.now ?? (() => new Date());
    if (o.source) this.source = o.source;
    else if (o.path) this.source = fileSource(o.path, { debounceMs: o.debounceMs, statPollMs: o.statPollMs, isOpen: o.isOpen });
    else throw new Error("DocWatcher: path or source is required");
  }

  /** 文書の場所 (ファイルのパス、または URL) */
  get path(): string {
    return this.source.location;
  }

  /** 監視を始める。前回の記録があれば、監視していない間の変化に追いつく */
  async start(): Promise<WatchUpdate> {
    const { bytes, mtime } = await this.source.read();
    const sha256 = sha256Of(bytes);
    this.archive = WatchArchive.open(this.o.recordsRoot, this.path, { docId: await docIdOf(bytes), sha256, name: this.source.name });
    const info = this.archive.info;
    const prevEnd = info.periods[info.periods.length - 1];
    info.periods.push({ start: this.now().toISOString() });
    this.archive.saveInfo();
    const known = info.lastSha256 !== undefined;
    const u = await this.process(bytes, mtime, known ? "catchUp" : "initial", true);
    if (u.gap) {
      info.gaps.push({ since: prevEnd?.end ?? info.lastSeenAt, detectedAt: this.now().toISOString(), sha256 });
      this.archive.saveInfo();
    }
    const onError = (err: unknown) => this.o.onError?.(err);
    this.unsubscribe = this.source.subscribe(() => {
      if (this.stopped) return;
      void this.serial(async () => {
        let r: { bytes: Uint8Array; mtime: Date };
        try {
          r = await this.source.read();
        } catch {
          // 消えた・読めない (名前を変えている途中など): 次の変化を待つ
          return;
        }
        await this.process(r.bytes, r.mtime, "saved");
      });
    }, onError);
    if (this.source.canWrite) {
      this.closeTimer = setInterval(() => void this.checkClosed(), this.o.closePollMs ?? 3000);
      void this.checkClosed();
    }
    return u;
  }

  /** 監視をやめる (監視した期間の終わりを記録する) */
  stop(): void {
    if (this.stopped) return;
    this.stopped = true;
    this.unsubscribe?.();
    if (this.closeTimer) clearInterval(this.closeTimer);
    if (this.archive) {
      const last = this.archive.info.periods[this.archive.info.periods.length - 1];
      if (last && !last.end) last.end = this.now().toISOString();
      this.archive.saveInfo();
    }
  }

  /** いまのコメントの状態 */
  comments(): CommentState[] {
    return [...this.archive.commentState().values()].sort((a, b) => a.seq - b.seq);
  }

  /** 処理を1つずつ順に行う */
  private serial(fn: () => Promise<unknown>): Promise<void> {
    const p = this.busy.then(fn).catch((err) => this.o.onError?.(err));
    this.busy = p;
    return p.then(() => undefined);
  }

  /** 読んだ文書を処理する (版の保管・コメントの記録・返信の待ち行列)。同じ内容なら何もしない */
  private async process(bytes: Uint8Array, mtime: Date, kind: WatchUpdateKind, always = false): Promise<WatchUpdate> {
    const a = this.archive;
    const sha256 = sha256Of(bytes);
    const now = this.now();
    const unchanged = a.info.lastSha256 === sha256;
    const toolWrite = a.info.toolWrites.some((w) => w.sha256 === sha256);
    const base: WatchUpdate = { kind, path: this.path, bytes, sha256, mtime, events: [], comments: [], gap: false };
    if ((unchanged || toolWrite) && !always) {
      a.info.lastSha256 = sha256;
      a.saveInfo();
      return { ...base, comments: this.comments() };
    }
    let version: SnapshotEntry | undefined;
    let events: CommentEvent[] = [];
    if (!toolWrite) {
      version = await a.addVersion(bytes, mtime, now);
      events = a.appendEvents(diffComments(a.commentState(), await extractComments(bytes), now, sha256));
      // 新しく見たコメントへの返信を待ち行列に入れる (書き込める読み取り元だけ)
      const seen = events.filter((e) => e.type === "seen");
      if (seen.length && this.source.canWrite) {
        const q = a.receipts();
        const queued = new Set([...q.pending, ...q.written, ...q.skipped].map((r) => r.key));
        for (const e of seen) if (!queued.has(e.key)) q.pending.push({ key: e.key, seq: e.seq, queuedAt: now.toISOString() });
        a.saveReceipts(q);
      }
    }
    const gap = kind === "catchUp" && !unchanged && !toolWrite;
    a.info.lastSha256 = sha256;
    a.info.lastSeenAt = now.toISOString();
    a.saveInfo();
    const u: WatchUpdate = { ...base, version, events, comments: this.comments(), gap };
    if (kind !== "initial" && kind !== "catchUp") this.o.onUpdate(u);
    return u;
  }

  /** 文書が閉じられていれば、待ち行列の返信を書き込む */
  checkClosed(): Promise<void> {
    return this.serial(async () => {
      if (this.stopped || !this.archive || !this.source.canWrite) return;
      const q = this.archive.receipts();
      if (q.pending.length === 0) return;
      if (await this.source.isOpen?.()) return;
      await this.writeReceipts();
    });
  }

  private async writeReceipts(): Promise<void> {
    const a = this.archive;
    let { bytes, mtime } = await this.source.read();
    // 最後に読んだときから変わっていれば、先に記録する
    if (sha256Of(bytes) !== a.info.lastSha256) {
      await this.process(bytes, mtime, "saved");
      ({ bytes, mtime } = await this.source.read());
    }
    const q = a.receipts();
    if (q.pending.length === 0) return;
    const now = this.now();
    const r = await addCommentReplies(
      bytes,
      q.pending.map((p) => ({ key: p.key, text: this.o.receiptText(p.seq, now) })),
      { now }
    );
    const seqOf = new Map(q.pending.map((p) => [p.key, p.seq]));
    const written = r.written.map((key) => ({ key, seq: seqOf.get(key)! }));
    const skipped = r.skipped.map((s) => ({ key: s.key, seq: seqOf.get(s.key)!, reason: s.reason }));
    let sha256 = a.info.lastSha256 ?? "";
    if (written.length) {
      sha256 = sha256Of(r.bytes);
      // 書き込みによる変化を新しい版として扱わないよう、先に記録する
      a.info.toolWrites.push({ sha256, at: now.toISOString() });
      a.saveInfo();
      // 書き込む直前に、文書が変わっていないか・開かれていないかを確かめる
      const latest = await this.source.read();
      if (sha256Of(latest.bytes) !== sha256Of(bytes) || (await this.source.isOpen?.())) return;
      await this.source.write!(r.bytes);
      a.info.lastSha256 = sha256;
      a.saveInfo();
      a.appendEvents(written.map((w) => ({ at: now.toISOString(), type: "receipt" as const, key: w.key, sha256 })));
    }
    q.pending = q.pending.filter((p) => !r.written.includes(p.key) && !r.skipped.some((s) => s.key === p.key));
    q.written.push(...written.map((w) => ({ ...w, at: now.toISOString(), sha256 })));
    q.skipped.push(...skipped.map((s) => ({ ...s, at: now.toISOString() })));
    a.saveReceipts(q);
    this.o.onReceipts?.({ path: this.path, written, skipped, sha256 });
  }
}
