/**
 * 拡張による外部送信の仲介 (メインプロセス)。
 *
 *  - 拡張の net.post はすべてここを通る。宛先・時刻・本文 (JSON)・結果を拡張ごとの送信履歴に記録する。
 *  - 送信前の確認: 拡張が求めた場合 (net.post の confirm)、または利用者の設定 (既定は確認する) で、
 *    宛先と本文を見せて送信するか拒否するかを尋ねる。
 *  - manifest で宣言していないオリジンへの送信は拒否し、拒否したことも記録する。
 *  - オフライン (または接続できない) ときは送信をキューに入れ、つながったら送る。
 *    拡張を無効にしたときは、その拡張のキューを破棄する。
 */

import { net, session } from "electron";
import * as fs from "fs";
import * as path from "path";
import type { ExtensionManifest, NetResponse } from "./extension-api";

export type SendStatus = "sent" | "failed" | "rejected" | "blocked" | "queued" | "discarded" | "opened";

export interface SendLogEntry {
  /** 記録の通し番号 */
  seq: number;
  /** 記録した時刻 (ISO 8601) */
  time: string;
  url: string;
  status: SendStatus;
  /** 送信した本文 (JSON。大きい場合は切り詰める) */
  body: string;
  truncated: boolean;
  httpStatus?: number;
  /** 送信前に利用者が確認したか */
  confirmed?: boolean;
  error?: string;
  /** キューから送ったものなら、キューに入れた時刻 */
  queuedAt?: string;
}

interface QueueItem {
  seq: number;
  url: string;
  body: string;
  queuedAt: string;
}

const MAX_LOG_ENTRIES = 500;
const MAX_LOGGED_BODY = 64 * 1024;
const MAX_BODY = 1024 * 1024;
const MAX_RESPONSE = 1024 * 1024;
const FLUSH_INTERVAL_MS = 60_000;

export interface MediatorOptions {
  /** 拡張ごとのデータの置き場所 (<dir>/<id>/send-log.json, send-queue.json) */
  dir: string;
  /** 利用者の設定で、送信前に確認するか */
  confirmByDefault(extId: string): boolean;
  /** 宛先と本文を見せて、送信してよいかを尋ねる */
  askUser(ext: ExtensionManifest, url: string, bodyJson: string): Promise<boolean>;
  /** 図の注釈のリンクを開いてよいかを尋ねる (URL 全体を見せる) */
  askOpen(ext: ExtensionManifest, url: string): Promise<boolean>;
  /** URL を既定のブラウザで開く */
  openExternal(url: string): Promise<void>;
  /** 送信履歴が変わったとき (設定画面の表示を更新するため) */
  onChange?(extId: string): void;
}

export class NetMediator {
  private seq = Date.now();
  private timer: NodeJS.Timeout;

  constructor(private readonly o: MediatorOptions) {
    this.timer = setInterval(() => void this.flushAll(), FLUSH_INTERVAL_MS);
    this.timer.unref?.();
  }

  private file(extId: string, name: string): string {
    return path.join(this.o.dir, extId, name);
  }

  private read<T>(extId: string, name: string, fallback: T): T {
    try {
      return JSON.parse(fs.readFileSync(this.file(extId, name), "utf-8")) as T;
    } catch {
      return fallback;
    }
  }

  private write(extId: string, name: string, value: unknown): void {
    const f = this.file(extId, name);
    fs.mkdirSync(path.dirname(f), { recursive: true });
    fs.writeFileSync(f, JSON.stringify(value), "utf-8");
  }

  log(extId: string): SendLogEntry[] {
    return this.read<SendLogEntry[]>(extId, "send-log.json", []);
  }

  queue(extId: string): QueueItem[] {
    return this.read<QueueItem[]>(extId, "send-queue.json", []);
  }

  private record(extId: string, e: Omit<SendLogEntry, "seq" | "time" | "truncated" | "body"> & { body: string; seq?: number }): void {
    const entries = this.log(extId);
    entries.push({
      ...e,
      seq: e.seq ?? this.seq++,
      time: new Date().toISOString(),
      body: e.body.length > MAX_LOGGED_BODY ? e.body.slice(0, MAX_LOGGED_BODY) : e.body,
      truncated: e.body.length > MAX_LOGGED_BODY,
    });
    this.write(extId, "send-log.json", entries.slice(-MAX_LOG_ENTRIES));
    this.o.onChange?.(extId);
  }

  /**
   * 拡張の net.post。allowed は宛先が manifest で宣言されたオリジンか (呼び出し側で判定する)。
   * 拒否・ブロックは例外、キューに入れた場合は queued: true の応答を返す。
   */
  async post(ext: ExtensionManifest, url: string, body: unknown, opts: { confirm?: boolean; allowed: boolean }): Promise<NetResponse> {
    let json: string;
    try {
      json = JSON.stringify(body ?? null);
    } catch {
      throw new Error("the body must be JSON-serializable");
    }
    if (json.length > MAX_BODY) throw new Error(`the body is limited to ${MAX_BODY} bytes`);
    if (!opts.allowed) {
      this.record(ext.id, { url, status: "blocked", body: json, error: "the destination is not declared in manifest.json" });
      throw new Error(`sending to ${safeOrigin(url)} is not allowed (not declared in manifest.json)`);
    }
    const mustConfirm = opts.confirm === true || this.o.confirmByDefault(ext.id);
    if (mustConfirm) {
      const ok = await this.o.askUser(ext, url, json);
      if (!ok) {
        this.record(ext.id, { url, status: "rejected", body: json, confirmed: false });
        throw new Error("the user declined to send this data");
      }
    }
    const seq = this.seq++;
    if (!net.isOnline()) return this.enqueue(ext.id, { seq, url, body: json, queuedAt: new Date().toISOString() }, mustConfirm);
    try {
      const res = await send(url, json);
      this.record(ext.id, { seq, url, status: res.ok ? "sent" : "failed", body: json, httpStatus: res.status, confirmed: mustConfirm || undefined });
      return res;
    } catch (err) {
      // 接続できない (オフライン・名前解決できない等) ときはキューに入れる。それ以外の失敗は記録して例外にする
      if (isConnectivityError(err)) {
        return this.enqueue(ext.id, { seq, url, body: json, queuedAt: new Date().toISOString() }, mustConfirm, err);
      }
      const message = err instanceof Error ? err.message : String(err);
      this.record(ext.id, { seq, url, status: "failed", body: json, confirmed: mustConfirm || undefined, error: message });
      throw new Error(`sending failed: ${message}`);
    }
  }

  private enqueue(extId: string, item: QueueItem, confirmed: boolean, err?: unknown): NetResponse {
    this.write(extId, "send-queue.json", [...this.queue(extId), item]);
    this.record(extId, {
      seq: item.seq,
      url: item.url,
      status: "queued",
      body: item.body,
      confirmed: confirmed || undefined,
      error: err ? (err instanceof Error ? err.message : String(err)) : "offline",
    });
    return { ok: false, status: 0, body: "", queued: true };
  }

  /** 拡張のキューを送る (オンラインのとき)。送れなかったものは残す */
  async flush(extId: string): Promise<void> {
    if (!net.isOnline()) return;
    const rest: QueueItem[] = [];
    for (const item of this.queue(extId)) {
      try {
        const res = await send(item.url, item.body);
        this.record(extId, { url: item.url, status: res.ok ? "sent" : "failed", body: item.body, httpStatus: res.status, queuedAt: item.queuedAt });
      } catch (err) {
        if (isConnectivityError(err)) rest.push(item);
        else this.record(extId, { url: item.url, status: "failed", body: item.body, queuedAt: item.queuedAt, error: String(err) });
      }
    }
    this.write(extId, "send-queue.json", rest);
  }

  private async flushAll(): Promise<void> {
    let ids: string[] = [];
    try {
      ids = fs.readdirSync(this.o.dir);
    } catch {
      return;
    }
    for (const id of ids) if (this.queue(id).length) await this.flush(id);
  }

  /**
   * 図の注釈のリンクを開く。URL に文書の内容を含めて外へ出すこともできるため、送信と同じく、
   * 利用者の設定 (既定は確認する) で URL を見せて確認し、送信履歴に記録する。allowed は宣言したオリジンか
   */
  async openLink(ext: ExtensionManifest, url: string, allowed: boolean): Promise<boolean> {
    if (!allowed) {
      this.record(ext.id, { url, status: "blocked", body: "", error: "the link's origin is not declared in manifest.json" });
      return false;
    }
    const mustConfirm = this.o.confirmByDefault(ext.id);
    if (mustConfirm && !(await this.o.askOpen(ext, url))) {
      this.record(ext.id, { url, status: "rejected", body: "", confirmed: false });
      return false;
    }
    await this.o.openExternal(url);
    this.record(ext.id, { url, status: "opened", body: "", confirmed: mustConfirm || undefined });
    return true;
  }

  /** 拡張を無効にしたとき: キューを破棄し、破棄したことを記録する */
  discard(extId: string): void {
    const q = this.queue(extId);
    if (q.length === 0) return;
    this.write(extId, "send-queue.json", []);
    for (const item of q) this.record(extId, { url: item.url, status: "discarded", body: item.body, queuedAt: item.queuedAt });
  }

  dispose(): void {
    clearInterval(this.timer);
  }
}

/** つながらないことによる失敗か (このときだけキューに入れて、あとで送り直す) */
function isConnectivityError(err: unknown): boolean {
  const msg = err instanceof Error ? err.message : String(err);
  return /ERR_(INTERNET_DISCONNECTED|NAME_NOT_RESOLVED|NAME_RESOLUTION_FAILED|CONNECTION_(REFUSED|RESET|CLOSED|TIMED_OUT|FAILED|ABORTED)|ADDRESS_UNREACHABLE|NETWORK_CHANGED|NETWORK_IO_SUSPENDED|TIMED_OUT|PROXY_CONNECTION_FAILED)/.test(
    msg
  );
}

function safeOrigin(url: string): string {
  try {
    return new URL(url).origin;
  } catch {
    return url;
  }
}

/** 送信専用のセッション (Cookie 等を持たない) から JSON を POST する */
async function send(url: string, json: string): Promise<NetResponse> {
  const res = await session.fromPartition("dra-ext-net").fetch(url, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: json,
  });
  const text = await res.text();
  return { ok: res.ok, status: res.status, body: text.slice(0, MAX_RESPONSE) };
}
