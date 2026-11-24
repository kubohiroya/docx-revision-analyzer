// Package one platform's CLI and, optionally, the Apple Silicon droplets.
import * as fs from "node:fs";
import * as path from "node:path";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const os = { darwin: "mac", win32: "win", linux: "linux" }[process.platform];
const args = process.argv.slice(2);
const platform = args.find((arg) => !arg.startsWith("--")) ?? `${os}-${process.arch}`;
if (!/^(linux|mac|win)-(x64|arm64)$/.test(platform) || args.some((arg) => arg.startsWith("--") && arg !== "--include-droplets")) {
  throw new Error("Usage: package-cli.mjs [<linux|mac|win>-<x64|arm64>] [--include-droplets]");
}
const [buildOs, arch] = platform.split("-");
const publicOs = { linux: "linux", mac: "macos", win: "windows" }[buildOs];
const output = path.join(root, "release", "assets");
fs.mkdirSync(output, { recursive: true });

async function writeArchive(name, zip) {
  const bytes = await zip.generateAsync({ type: "nodebuffer", platform: "UNIX", compression: "DEFLATE", compressionOptions: { level: 9 } });
  const digest = createHash("sha256").update(bytes).digest("hex");
  fs.writeFileSync(path.join(output, name), bytes);
  fs.writeFileSync(path.join(output, `${name}.sha256`), `${digest}  ${name}\n`);
  console.log(`${name}: ${bytes.length} bytes, sha256 ${digest}`);
}

const executable = `docx-revision-analyzer${buildOs === "win" ? ".exe" : ""}`;
const cli = new JSZip();
cli.file(executable, fs.readFileSync(path.join(root, "release", "cli", platform, executable)), { unixPermissions: 0o100755 });
cli.file("LICENSE", fs.readFileSync(path.join(root, "LICENSE")), { unixPermissions: 0o100644 });
await writeArchive(`docx-revision-analyzer-${publicOs}-${arch}.zip`, cli);

if (args.includes("--include-droplets")) {
  if (platform !== "mac-arm64") throw new Error("Droplet release assets are built on mac-arm64.");
  for (const command of ["chart", "flow"]) {
    const name = `docx-revision-${command}.app`;
    const base = path.join(root, "release", "droplet", platform);
    const zip = new JSZip();
    function add(relative) {
      const full = path.join(base, relative);
      const stat = fs.lstatSync(full);
      const archivePath = relative.split(path.sep).join("/");
      if (stat.isSymbolicLink()) {
        zip.file(archivePath, fs.readlinkSync(full), { unixPermissions: 0o120777 });
      } else if (stat.isDirectory()) {
        zip.file(`${archivePath}/`, null, { dir: true, unixPermissions: stat.mode });
        for (const entry of fs.readdirSync(full).sort()) add(path.join(relative, entry));
      } else {
        zip.file(archivePath, fs.readFileSync(full), { unixPermissions: stat.mode });
      }
    }
    add(name);
    await writeArchive(`docx-revision-${command}-macos-arm64.app.zip`, zip);
  }
}
