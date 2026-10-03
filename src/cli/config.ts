/**
 * config.ts
 *
 * 実行ファイルと同じフォルダに置いた設定ファイル (<ツール名>.yml / .yaml) から、
 * オプションの既定値を読み込む。
 *
 *   # docx-revision-flow.yml の例
 *   gap-threshold: 2
 *   bulk-chars: 200
 *   title: 研究計画調書の編集フロー
 *
 * キーは各ツールのオプションの長い名前 (先頭の -- を除いたもの)。
 * 優先順位は コマンドラインの指定 > 設定ファイル > 組み込みの既定値。
 * output / rules / template の相対パスは設定ファイルのあるフォルダを基準にする
 * (ドラッグ&ドロップ起動では作業フォルダが定まらないため)。
 *
 * 探す場所 (最初に見つかったものを使う):
 *   1. 単体バイナリの場合: 実行ファイルのフォルダ。macOS のドロップレット (.app の中) の場合は、
 *      .app と同じフォルダを先に探す
 *   2. Node.js で実行している場合: 起動したコマンド (npm の bin に置かれたリンク等) のフォルダ
 */

import * as fs from "fs";
import * as path from "path";
import { parse } from "yaml";
import { Command, Option } from "commander";
import { t } from "../lib/i18n";

export interface ToolConfig {
  path: string;
  dir: string;
  values: Record<string, unknown>;
}

/** 設定ファイルを探すフォルダの候補 */
export function configCandidateDirs(): string[] {
  const dirs: string[] = [];
  const exe = process.execPath;
  // node / bun 本体で実行している場合は、実行ファイルのフォルダは対象外 (ランタイムのフォルダのため)
  if (!/^(node|nodejs|bun)(\.exe)?$/i.test(path.basename(exe))) {
    const exeDir = path.dirname(exe);
    const app = exeDir.match(/^(.*)[/\\][^/\\]+\.app[/\\]Contents[/\\](?:Resources|MacOS)$/);
    if (app) dirs.push(app[1]);
    dirs.push(exeDir);
  }
  if (process.argv[1]) dirs.push(path.dirname(path.resolve(process.argv[1])));
  return [...new Set(dirs)];
}

/** 設定ファイルを探して読み込む。見つからなければ config も error も undefined */
export function loadToolConfig(tool: string): { config?: ToolConfig; error?: string } {
  for (const dir of configCandidateDirs()) {
    for (const name of [`${tool}.yml`, `${tool}.yaml`]) {
      const file = path.join(dir, name);
      if (!fs.existsSync(file)) continue;
      try {
        const values = parse(fs.readFileSync(file, "utf-8")) ?? {};
        if (typeof values !== "object" || Array.isArray(values)) {
          return { error: t("configInvalid", file, t("configNotMapping")) };
        }
        return { config: { path: file, dir, values: values as Record<string, unknown> } };
      } catch (err) {
        return { error: t("configInvalid", file, err instanceof Error ? err.message : String(err)) };
      }
    }
  }
  return {};
}

/**
 * コマンドラインで指定されなかったオプションに、設定ファイルの値を入れる。
 * 実行前フックで行うため、.action() に渡るオプションに反映される。
 */
export function applyToolConfig(program: Command, config: ToolConfig): void {
  program.hook("preAction", (cmd) => {
    console.error(t("configLoaded", config.path));
    for (const [key, raw] of Object.entries(config.values)) {
      const opt: Option | undefined = cmd.options.find((o) => o.long === `--${key}`);
      if (!opt) {
        console.error(t("warning", t("configUnknownKey", key, config.path)));
        continue;
      }
      const name = opt.attributeName();
      if (cmd.getOptionValueSource(name) === "cli" || raw === null || raw === undefined) continue;

      let value: unknown;
      if (opt.isBoolean()) {
        value = raw === true || String(raw).toLowerCase() === "true";
      } else if (key === "output") {
        // 複数ファイルを処理するときは、出力先を1つに決められないため使わない
        if (cmd.args.length > 1) {
          console.error(t("warning", t("configOutputIgnored", config.path)));
          continue;
        }
        value = path.resolve(config.dir, String(raw));
      } else if (key === "rules" || key === "template") {
        value = path.resolve(config.dir, String(raw));
      } else {
        value = String(raw);
      }
      if (opt.argChoices && !opt.argChoices.includes(String(value))) {
        console.error(t("warning", t("configBadValue", key, String(raw), config.path)));
        continue;
      }
      cmd.setOptionValueWithSource(name, value, "config");
    }
  });
}

/** CLI の起動時に使う: 設定ファイルを読み、読めなければエラーを表示して終了する */
export function loadToolConfigOrExit(tool: string, loaded: { config?: ToolConfig; error?: string }): ToolConfig | undefined {
  if (loaded.error) {
    console.error(t("errorFor", tool, loaded.error));
    process.exit(1);
  }
  return loaded.config;
}
