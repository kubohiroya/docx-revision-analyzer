/**
 * trackLock.ts
 *
 * テンプレートに Word の「変更履歴のロック」をかける (word/settings.xml に
 * <w:documentProtection w:edit="trackedChanges" w:enforcement="1" .../> を書き込む)。
 * あわせて「変更履歴の記録」をオンにし、保存時に作成者・日時を削除する設定を外す。
 *
 * パスワードのハッシュ値は、Word と同じ方式で作る (ECMA-376 Part 4 の documentProtection と、
 * Apache POI の XWPFSettings.setEnforcementEditValue / CryptoFunctions を参照。ただし 1. のバイトの扱いは
 * POI と違い、Word が書き出したハッシュ値に合わせてある):
 *  1. パスワード (先頭15文字) から、Word 2003 以前の 32 ビットの鍵を作る
 *  2. その鍵のバイト順を逆にして 16 進の文字列 (大文字) にし、UTF-16LE のバイト列にする
 *  3. salt + 2. を SHA-512 でハッシュし、続けて「前回のハッシュ値 + 回数 (32 ビット LE)」を spinCount 回ハッシュする
 *
 * ファイルシステムは使わない (ファイルへの保存は node/lockFile.ts)。ハッシュ関数と乱数は、既定では
 * Web Crypto (globalThis.crypto) を使い、Node.js では node/lockFile.ts が node:crypto のものを渡す。
 */

import JSZip from "jszip";
import { t } from "./i18n";
import type { DocxInput } from "./input";
import {
  ELEMENTS_AFTER_TRACK_REVISIONS,
  ensureTrackRevisions,
  insertSettingsElement,
  parseHistorySettings,
  removeSettingsElement,
  wmlPrefix,
} from "./historySettings";

const SETTINGS_PART = "word/settings.xml";
/** Word 2013 以降と同じ: SHA-512 (cryptAlgorithmSid 14)、100000 回 */
const DEFAULT_SPIN_COUNT = 100000;

export type Digest = (data: Uint8Array) => Uint8Array | Promise<Uint8Array>;

export interface LockOptions {
  /** ロックを外すためのパスワード。空・省略ならパスワード無し (誰でも Word のメニューから外せる) */
  password?: string;
  /** SHA-512 の関数 (既定: Web Crypto) */
  digest?: Digest;
  /** 16 バイトの salt (既定: Web Crypto の乱数) */
  salt?: Uint8Array;
  spinCount?: number;
}

export interface LockResult {
  output: Uint8Array;
  /** もともとロックされていたか (パスワードは置き換わる) */
  wasLocked: boolean;
  /** もともと「変更履歴の記録」がオンだったか */
  wasTracking: boolean;
  /** 保存時に作成者・日時を削除する設定を外したか */
  removedPersonalInfoSetting: boolean;
  /** パスワードを付けたか */
  withPassword: boolean;
}

// Word 2003 以前のパスワードの鍵 (MS-OFFCRYPTO 2.3.7.4 Binary Document Password Verifier Derivation Method 2)
const INITIAL_CODE_ARRAY = [
  0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3,
];

const ENCRYPTION_MATRIX = [
  [0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09],
  [0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf],
  [0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0],
  [0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40],
  [0xd849, 0xa0b3, 0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5],
  [0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a],
  [0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9],
  [0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0],
  [0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec],
  [0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10],
  [0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168],
  [0x76b4, 0xed68, 0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c],
  [0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad],
  [0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc],
  [0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4],
];

const MAX_PASSWORD_LENGTH = 15;

/** 各文字の下位バイト (0 なら上位バイト) */
function ansiBytes(password: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i);
    const low = c & 0xff;
    out.push(low !== 0 ? low : (c >>> 8) & 0xff);
  }
  return out;
}

function rotateLeft15(v: number): number {
  return ((v & 0x4000) === 0 ? 0 : 1) | ((v << 1) & 0x7fff);
}

/**
 * Word 2003 以前のパスワードの鍵を、バイト順を逆にした 16 進の文字列で返す
 * (例: 鍵 0x64CEED7E → "7EEDCE64")。空のパスワードは "00000000"
 */
export function legacyPasswordKey(password: string): string {
  let key = 0;
  if (password.length > 0) {
    const bytes = ansiBytes(password.slice(0, MAX_PASSWORD_LENGTH));
    let high = INITIAL_CODE_ARRAY[bytes.length - 1];
    let line = MAX_PASSWORD_LENGTH - bytes.length;
    for (const b of bytes) {
      const row = ENCRYPTION_MATRIX[line++];
      for (let bit = 0; bit < 7; bit++) if (b & (1 << bit)) high ^= row[bit];
    }
    // 下位の語。バイトは符号なしとして XOR する (Apache POI は Java の byte のまま符号付きで XOR するため、
    // 0x80 以上のバイトを含むパスワード (日本語など) で Word と食い違う。Word for Mac で確かめた)
    let low = 0;
    for (let i = bytes.length - 1; i >= 0; i--) {
      low = rotateLeft15(low);
      low = (low ^ bytes[i]) & 0xffff;
    }
    low = rotateLeft15(low);
    low = (low ^ bytes.length ^ 0xce4b) & 0xffff;
    key = ((high & 0xffff) << 16) | low;
  }
  const hex = (n: number) => (n & 0xff).toString(16).toUpperCase().padStart(2, "0");
  return hex(key) + hex(key >>> 8) + hex(key >>> 16) + hex(key >>> 24);
}

function utf16le(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = c >>> 8;
  }
  return out;
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a, 0);
  out.set(b, a.length);
  return out;
}

function webCrypto(): { subtle?: { digest(alg: string, data: Uint8Array): Promise<ArrayBuffer> }; getRandomValues?(a: Uint8Array): Uint8Array } {
  return ((globalThis as { crypto?: unknown }).crypto ?? {}) as ReturnType<typeof webCrypto>;
}

const webDigest: Digest = async (data) => {
  const subtle = webCrypto().subtle;
  if (!subtle) throw new Error(t("errNoCrypto"));
  return new Uint8Array(await subtle.digest("SHA-512", data));
};

function randomSalt(): Uint8Array {
  const c = webCrypto();
  if (!c.getRandomValues) throw new Error(t("errNoCrypto"));
  return c.getRandomValues(new Uint8Array(16));
}

/** documentProtection のハッシュ値 (SHA-512) を求める */
export async function hashLockPassword(
  password: string,
  salt: Uint8Array,
  spinCount = DEFAULT_SPIN_COUNT,
  digest: Digest = webDigest
): Promise<Uint8Array> {
  let hash = await digest(concat(salt, utf16le(legacyPasswordKey(password))));
  const iterator = new Uint8Array(4);
  for (let i = 0; i < spinCount; i++) {
    iterator[0] = i & 0xff;
    iterator[1] = (i >>> 8) & 0xff;
    iterator[2] = (i >>> 16) & 0xff;
    iterator[3] = (i >>> 24) & 0xff;
    hash = await digest(concat(hash, iterator));
  }
  return hash;
}

function base64(bytes: Uint8Array): string {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s);
}

function fromBase64(s: string): Uint8Array {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** CT_Settings のスキーマ上、w:documentProtection より後ろに置かれる要素 */
const ELEMENTS_AFTER_PROTECTION = new Set(
  [...ELEMENTS_AFTER_TRACK_REVISIONS].filter((n) => !["doNotTrackMoves", "doNotTrackFormatting", "documentProtection"].includes(n))
);

/** settings.xml の文字列に、変更履歴のロックを書き込む */
export async function lockSettingsXml(xml: string, opts: LockOptions = {}): Promise<string> {
  const w = wmlPrefix(xml);
  let out = xml;
  for (const name of ["documentProtection", "removePersonalInformation", "removeDateAndTime"]) out = removeSettingsElement(out, w, name);
  out = ensureTrackRevisions(out, w);
  let attrs = `${w}:edit="trackedChanges" ${w}:enforcement="1"`;
  if (opts.password) {
    const salt = opts.salt ?? randomSalt();
    const spin = opts.spinCount ?? DEFAULT_SPIN_COUNT;
    const hash = await hashLockPassword(opts.password, salt, spin, opts.digest ?? webDigest);
    attrs +=
      ` ${w}:cryptProviderType="rsaAES" ${w}:cryptAlgorithmClass="hash" ${w}:cryptAlgorithmType="typeAny"` +
      ` ${w}:cryptAlgorithmSid="14" ${w}:cryptSpinCount="${spin}" ${w}:hash="${base64(hash)}" ${w}:salt="${base64(salt)}"`;
  }
  return insertSettingsElement(out, w, `<${w}:documentProtection ${attrs}/>`, ELEMENTS_AFTER_PROTECTION);
}

function readProtectionAttrs(settingsXml: string | undefined): Record<string, string> | undefined {
  const m = settingsXml?.match(/<(?:\w+:)?documentProtection\b([^>]*)\/?>/);
  if (!m) return undefined;
  const a: Record<string, string> = {};
  for (const x of m[1].matchAll(/(?:\w+:)?(\w+)="([^"]*)"/g)) a[x[1]] = x[2];
  return a;
}

/** docx が変更履歴のロックをかけた状態か */
export async function isTrackChangesLocked(input: DocxInput): Promise<boolean> {
  const zip = await JSZip.loadAsync(input);
  const a = readProtectionAttrs(await zip.file(SETTINGS_PART)?.async("string"));
  return !!a && a.edit === "trackedChanges" && ["1", "true", "on"].includes((a.enforcement ?? "").toLowerCase());
}

/** docx に、パスワード付きの変更履歴のロックがかかっているか */
export async function hasLockPassword(input: DocxInput): Promise<boolean> {
  const zip = await JSZip.loadAsync(input);
  const a = readProtectionAttrs(await zip.file(SETTINGS_PART)?.async("string"));
  return !!a && a.edit === "trackedChanges" && !!(a.hash ?? a.hashValue);
}

/**
 * ロックのパスワードが合っているかを確かめる (SHA-512 でハッシュしたものだけ。ロックが無ければ false)。
 * パスワード無しのロックには、空のパスワードで true
 */
export async function verifyLockPassword(input: DocxInput, password: string, digest: Digest = webDigest): Promise<boolean> {
  const zip = await JSZip.loadAsync(input);
  const a = readProtectionAttrs(await zip.file(SETTINGS_PART)?.async("string"));
  if (!a || a.edit !== "trackedChanges") return false;
  const hash = a.hash ?? a.hashValue;
  const salt = a.salt ?? a.saltValue;
  if (!hash || !salt) return password === "";
  if ((a.cryptAlgorithmSid ?? "14") !== "14" && (a.algorithmName ?? "SHA-512") !== "SHA-512") return false;
  const spin = Number(a.cryptSpinCount ?? a.spinCount ?? DEFAULT_SPIN_COUNT);
  return base64(await hashLockPassword(password, fromBase64(salt), spin, digest)) === hash;
}

/** docx のバイト列に変更履歴のロックをかけた docx を返す */
export async function lockTrackChangesInDocx(input: DocxInput, opts: LockOptions = {}): Promise<LockResult> {
  const zip = await JSZip.loadAsync(input);
  const xml = await zip.file(SETTINGS_PART)?.async("string");
  if (xml === undefined) throw new Error(t("errLockNoSettings"));
  const before = parseHistorySettings(xml);
  const prev = readProtectionAttrs(xml);
  const wasLocked = !!prev && prev.edit === "trackedChanges" && ["1", "true", "on"].includes((prev.enforcement ?? "").toLowerCase());
  const patched = await lockSettingsXml(xml, opts);
  const after = parseHistorySettings(patched);
  if (!after.trackRevisions || after.removePersonalInformation || after.removeDateAndTime) throw new Error(t("errPatchFailed"));
  zip.file(SETTINGS_PART, patched);
  return {
    output: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
    wasLocked,
    wasTracking: before.trackRevisions,
    removedPersonalInfoSetting: before.removePersonalInformation || before.removeDateAndTime,
    withPassword: !!opts.password,
  };
}
