// Exercise the public CLI in Node and, when supplied, as a compiled executable.
// DRA_CLI_BINARY=/absolute/path/to/docx-revision-analyzer npm run test:cli
import assert from "node:assert/strict";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-cli with spaces-"));
const binary = process.env.DRA_CLI_BINARY;
const cli = path.join(tmp, binary ? `docx-revision-analyzer${process.platform === "win32" ? ".exe" : ""}` : "cli.cjs");
const commands = ["chart", "flow", "score", "snapshot", "versions"];
const version = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version;

function run(args) {
  const result = spawnSync(binary ? cli : process.execPath, binary ? args : [cli, ...args], {
    cwd: tmp,
    encoding: "utf8",
    env: { ...process.env, LANG: "en_US.UTF-8", LC_ALL: "en_US.UTF-8" },
  });
  assert.ifError(result.error);
  return result;
}
function success(args) {
  const result = run(args);
  assert.equal(result.status, 0, `${args.join(" ")}\n${result.stderr}`);
  return result;
}
function failure(args) {
  const result = run(args);
  assert.notEqual(result.status, 0, args.join(" "));
  return result;
}
function config(tool, text) {
  fs.writeFileSync(path.join(tmp, `${tool}.yml`), text);
}

try {
  if (binary) {
    fs.copyFileSync(path.resolve(binary), cli);
    fs.chmodSync(cli, 0o755);
  } else {
    await esbuild.build({
      entryPoints: [path.join(root, "src/cli/index.ts")],
      bundle: true, platform: "node", format: "cjs", outfile: cli, logLevel: "silent",
    });
  }

  const help = success(["--lang", "en", "--help"]).stdout;
  for (const name of commands) {
    assert.match(help, new RegExp(`\\n  ${name}\\s`));
    const detailed = success([name, "--lang", "en", "--help"]).stdout;
    assert.match(detailed, new RegExp(`Usage: docx-revision-analyzer ${name} `));
    assert.equal(success(["help", name, "--lang", "en"]).stdout, detailed);
    assert.match(success(["--lang", "ja", name, "--help"]).stdout, /表示言語/);
    failure([name]);
    failure([name, "--unknown-option"]);
  }
  assert.match(success(["--lang=ja", "--help"]).stdout, /サブコマンド/);
  assert.equal(success(["--version"]).stdout.trim(), version);
  failure([]);
  assert.match(failure(["unknown"]).stderr, /unknown command/);
  assert.match(failure(["help", "unknown", "--lang", "en"]).stderr, /Unknown subcommand/);
  failure(["--lang", "invalid", "flow", "--help"]);
  failure(["flow", "--lang", "invalid", "--help"]);
  console.log("ok root/subcommand help, languages, version, and argument errors");

  const input = path.join(tmp, "document with spaces.docx");
  fs.copyFileSync(path.join(root, "fixtures/flow-demo.docx"), input);
  for (const name of ["chart", "flow"]) {
    const output = path.join(tmp, `${name}.svg`);
    success([name, input, "--lang", "en", "-o", output, "--json"]);
    assert.match(fs.readFileSync(output, "utf8"), /<svg\b/);
    assert.equal(JSON.parse(fs.readFileSync(output.replace(/\.svg$/, ".json"), "utf8")).tool, `docx-revision-${name}`);
    const mixed = failure([name, input, "missing.docx", "--lang", "en"]);
    assert.match(mixed.stderr, /missing\.docx/);
    failure([name, input, input, "-o", output, "--lang", "en"]);
    const leadingDash = `-${name}.docx`;
    fs.copyFileSync(input, path.join(tmp, leadingDash));
    success([name, "--lang", "en", "--", leadingDash]);
    assert.ok(fs.existsSync(path.join(tmp, `-${name}${name === "flow" ? "-flow" : ""}.svg`)));
  }
  const score = JSON.parse(success(["score", input, "--pretty", "--lang", "en"]).stdout);
  assert.equal(typeof score.score, "number");
  assert.equal(score.file, path.basename(input));
  failure(["score", "missing.docx", "--lang", "en"]);
  const outDir = path.join(tmp, "earlier versions");
  success(["versions", input, "--format", "txt", "--out-dir", outDir, "--lang", "en"]);
  assert.ok(fs.readdirSync(outDir).some((file) => file.endsWith(".versions.json")));
  const archive = path.join(tmp, "archive");
  success(["snapshot", archive, input, "--lang", "en"]);
  assert.ok(fs.existsSync(path.join(archive, "index.html")));
  console.log("ok five subcommands, output formats, spaces, --, and partial failures");

  config("docx-revision-chart", "broken: [\n");
  config("docx-revision-flow", "lang: ja\ntitle: Configured flow title\noutput: configured.svg\n");
  success(["--lang", "en", "--help"]);
  const configuredHelp = success(["flow", "--help"]).stdout;
  assert.match(configuredHelp, /表示言語/);
  assert.match(success(["--lang", "en", "flow", "--help"]).stdout, /Display language/);
  success(["flow", input]);
  assert.match(fs.readFileSync(path.join(tmp, "configured.svg"), "utf8"), /Configured flow title/);
  const override = path.join(tmp, "override.svg");
  success(["flow", input, "-o", override, "--title", "Explicit title", "--lang", "en"]);
  assert.match(fs.readFileSync(override, "utf8"), /Explicit title/);
  failure(["chart", input, "--lang", "en"]);
  console.log("ok selected settings only, legacy YAML names, relative paths, and CLI precedence");
} finally {
  fs.rmSync(tmp, { recursive: true, force: true });
}
