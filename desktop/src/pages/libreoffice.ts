/**
 * インストールされている LibreOffice で .docx を PDF にする (メインプロセス)。LibreOffice は同梱しない。
 *
 *  - 探す場所: 設定で指定したパス、macOS は /Applications/LibreOffice.app、Windows は Program Files、Linux は PATH
 *  - 専用の設定フォルダ (-env:UserInstallation) で動かし、利用者が開いている LibreOffice とぶつからないようにする
 *  - macOS の画面なしの LibreOffice は、文字の描画が fontconfig になり OS のフォントの場所を知らない (日本語が抜け落ちる)。
 *    FONTCONFIG_FILE で macOS のフォントの場所と、Mac に無い Windows のフォントの代用の規則を渡す (issue #59 の試作)
 *  - 画像の解像度は PDF の書き出しの設定で下げる (LibreOffice 7.4 以降の JSON の指定)
 *  - 変更履歴の印は表示しない設定 (w:revisionView) にしてから渡す (withoutRevisionMarks)
 *  - いつも1つずつ、優先度を下げて動かし、一定時間で終わらなければ止める (常駐はさせない)
 */
import { spawn } from "child_process";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { pathToFileURL } from "url";
import JSZip from "jszip";
import { ELEMENTS_AFTER_TRACK_REVISIONS, insertSettingsElement, removeSettingsElement, wmlPrefix } from "../../../src/lib/historySettings";

export type ImageQuality = "low" | "medium" | "high";

/** LibreOffice の soffice の場所 (見つからなければ undefined) */
export function findSoffice(configured?: string): string | undefined {
  const candidates: string[] = [];
  if (configured) candidates.push(configured);
  if (process.platform === "darwin") {
    candidates.push("/Applications/LibreOffice.app/Contents/MacOS/soffice", path.join(os.homedir(), "Applications/LibreOffice.app/Contents/MacOS/soffice"));
  } else if (process.platform === "win32") {
    for (const base of [process.env.ProgramFiles, process.env["ProgramFiles(x86)"]]) {
      if (base) candidates.push(path.join(base, "LibreOffice", "program", "soffice.exe"));
    }
  }
  for (const dir of (process.env.PATH ?? "").split(path.delimiter)) {
    if (dir) candidates.push(path.join(dir, process.platform === "win32" ? "soffice.exe" : "soffice"));
  }
  return candidates.find((p) => {
    try {
      return fs.statSync(p).isFile();
    } catch {
      return false;
    }
  });
}

/** Mac に無い Windows のフォント → macOS のフォント */
const FONT_ALIASES: [string[], string][] = [
  [["MS Mincho", "ＭＳ 明朝", "MS PMincho", "ＭＳ Ｐ明朝", "Yu Mincho", "游明朝", "YuMincho", "HG明朝E", "HGMinchoE"], "Hiragino Mincho ProN"],
  [["MS Gothic", "ＭＳ ゴシック", "MS PGothic", "ＭＳ Ｐゴシック", "MS UI Gothic", "Yu Gothic", "游ゴシック", "YuGothic", "Meiryo", "メイリオ", "Meiryo UI"], "Hiragino Sans"],
];

/** macOS 用の fontconfig の設定ファイルの中身 */
export function macFontconfig(libreOfficeFonts: string, cacheDir: string): string {
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const dirs = ["/System/Library/Fonts", "/System/Library/Fonts/Supplemental", "/Library/Fonts", "~/Library/Fonts", libreOfficeFonts];
  const aliases = FONT_ALIASES.flatMap(([from, to]) =>
    from.map((f) => `  <alias binding="same"><family>${esc(f)}</family><prefer><family>${esc(to)}</family></prefer></alias>`)
  );
  return [
    `<?xml version="1.0"?>`,
    `<!DOCTYPE fontconfig SYSTEM "fonts.dtd">`,
    `<fontconfig>`,
    ...dirs.map((d) => `  <dir>${esc(d)}</dir>`),
    `  <cachedir>${esc(cacheDir)}</cachedir>`,
    ...aliases,
    `</fontconfig>`,
    ``,
  ].join("\n");
}

/** PDF の書き出しの設定 (--convert-to に渡す JSON) */
export function exportFilter(quality: ImageQuality): string {
  if (quality === "high") return "pdf:writer_pdf_Export";
  const dpi = quality === "low" ? 96 : 150;
  const opts = {
    ReduceImageResolution: { type: "boolean", value: "true" },
    MaxImageResolution: { type: "long", value: String(dpi) },
    Quality: { type: "long", value: quality === "low" ? "60" : "75" },
    UseLosslessCompression: { type: "boolean", value: "false" },
  };
  return `pdf:writer_pdf_Export:${JSON.stringify(opts)}`;
}

const WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SETTINGS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const SETTINGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

/**
 * 変更履歴の印 (挿入の下線・削除した文字・書式の変更の印) を表示しない設定 (w:revisionView) にした .docx。
 * LibreOffice は文書の表示の設定のまま PDF にするので、そのままでは削除した文字も描かれ、改行・改ページの位置が
 * 変更を反映した後の文書 (Word の「シンプルな変更履歴」の表示や、flow の模式図) とずれる。
 * settings.xml の無い文書には、それを加える
 */
export async function withoutRevisionMarks(docx: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(docx);
  let xml = await zip.file("word/settings.xml")?.async("string");
  if (xml) {
    const w = wmlPrefix(xml);
    xml = insertSettingsElement(
      removeSettingsElement(xml, w, "revisionView"),
      w,
      `<${w}:revisionView ${w}:markup="0" ${w}:insDel="0" ${w}:formatting="0"/>`,
      new Set(["trackRevisions", ...ELEMENTS_AFTER_TRACK_REVISIONS])
    );
  } else {
    const types = await zip.file("[Content_Types].xml")?.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")?.async("string");
    if (!types || !rels) return docx;
    xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${WML}"><w:revisionView w:markup="0" w:insDel="0" w:formatting="0"/></w:settings>`;
    if (!types.includes(`PartName="/word/settings.xml"`)) {
      zip.file("[Content_Types].xml", types.replace("</Types>", `<Override PartName="/word/settings.xml" ContentType="${SETTINGS_TYPE}"/></Types>`));
    }
    zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", `<Relationship Id="rIdDraSettings" Type="${SETTINGS_REL}" Target="settings.xml"/></Relationships>`));
  }
  zip.file("word/settings.xml", xml);
  return zip.generateAsync({ type: "uint8array" });
}

export interface ConvertOptions {
  soffice: string;
  /** 専用の設定フォルダなどを置く場所 (アプリのデータフォルダの下) */
  workDir: string;
  quality: ImageQuality;
  timeoutMs?: number;
}

let queue: Promise<unknown> = Promise.resolve();

/** .docx を PDF にする (いつも1つずつ)。PDF のバイト列を返す */
export function convertDocxToPdf(docx: Uint8Array, o: ConvertOptions): Promise<Uint8Array> {
  const run = queue.then(() => convertNow(docx, o));
  queue = run.catch(() => undefined);
  return run;
}

async function convertNow(docx: Uint8Array, o: ConvertOptions): Promise<Uint8Array> {
  const profile = path.join(o.workDir, "libreoffice-profile");
  const tmp = await fs.promises.mkdtemp(path.join(o.workDir, "convert-"));
  try {
    const input = path.join(tmp, "document.docx");
    let source = docx;
    try {
      source = await withoutRevisionMarks(docx);
    } catch {
      // 設定を書き換えられない文書は、そのまま変換する (変更履歴の印が描かれる)
    }
    await fs.promises.writeFile(input, source);
    const env = { ...process.env };
    if (process.platform === "darwin") {
      const resources = path.resolve(path.dirname(o.soffice), "..", "Resources", "fonts");
      const conf = path.join(o.workDir, "fonts.conf");
      await fs.promises.writeFile(conf, macFontconfig(resources, path.join(o.workDir, "fontconfig-cache")));
      env.FONTCONFIG_FILE = conf;
    }
    const args = [
      `-env:UserInstallation=${pathToFileURL(profile).href}`,
      "--headless",
      "--norestore",
      "--nologo",
      "--nolockcheck",
      "--convert-to",
      exportFilter(o.quality),
      "--outdir",
      tmp,
      input,
    ];
    // 優先度を下げる (macOS / Linux は nice、Windows は起動後に「通常以下」)
    const useNice = process.platform !== "win32" && fs.existsSync("/usr/bin/nice");
    const child = useNice ? spawn("/usr/bin/nice", ["-n", "10", o.soffice, ...args], { env }) : spawn(o.soffice, args, { env, windowsHide: true });
    if (process.platform === "win32" && child.pid) {
      try {
        os.setPriority(child.pid, os.constants.priority.PRIORITY_BELOW_NORMAL);
      } catch {
        // 優先度を変えられなくても変換は続ける
      }
    }
    let stderr = "";
    child.stderr?.on("data", (d) => (stderr += String(d)));
    const code = await new Promise<number | null>((resolve, reject) => {
      const timer = setTimeout(() => {
        child.kill("SIGKILL");
        reject(new Error(`LibreOffice did not finish in ${Math.round((o.timeoutMs ?? 60_000) / 1000)} s`));
      }, o.timeoutMs ?? 60_000);
      child.on("error", (err) => {
        clearTimeout(timer);
        reject(err);
      });
      child.on("exit", (c) => {
        clearTimeout(timer);
        resolve(c);
      });
    });
    const out = path.join(tmp, "document.pdf");
    if (code !== 0 || !fs.existsSync(out)) throw new Error(`LibreOffice failed (exit ${code}): ${stderr.trim().slice(0, 300)}`);
    return new Uint8Array(await fs.promises.readFile(out));
  } finally {
    await fs.promises.rm(tmp, { recursive: true, force: true });
  }
}
