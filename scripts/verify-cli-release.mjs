// Validate release ZIP contents, executable architectures, and SHA-256 digests.
import assert from "node:assert/strict";
import * as fs from "node:fs";
import * as path from "node:path";
import { createHash } from "node:crypto";
import JSZip from "jszip";

const dir = path.resolve(process.argv[2] ?? "release/assets");
const complete = process.argv.includes("--complete");
const cliNames = ["linux", "macos", "windows"].flatMap((os) => ["x64", "arm64"].map((arch) => `docx-revision-analyzer-${os}-${arch}.zip`));
const dropletNames = ["chart", "flow"].map((command) => `docx-revision-${command}-macos-arm64.app.zip`);
const expected = [...cliNames, ...dropletNames].sort();
const names = fs.readdirSync(dir).filter((name) => name.endsWith(".zip")).sort();
assert.ok(names.length > 0, "No release ZIPs found");
if (complete) assert.deepEqual(names, expected, "The release must contain all six CLI ZIPs and both droplets");

function checkArchitecture(bytes, os, arch) {
  if (os === "linux") {
    assert.equal(bytes.subarray(0, 4).toString("hex"), "7f454c46");
    assert.equal(bytes[4], 2, "Expected a 64-bit ELF executable");
    assert.equal(bytes[5], 1, "Expected little-endian ELF");
    assert.equal(bytes.readUInt16LE(18), arch === "x64" ? 0x3e : 0xb7);
  } else if (os === "macos") {
    assert.equal(bytes.readUInt32LE(0), 0xfeedfacf, "Expected 64-bit Mach-O");
    assert.equal(bytes.readUInt32LE(4), arch === "x64" ? 0x01000007 : 0x0100000c);
  } else {
    assert.equal(bytes.subarray(0, 2).toString(), "MZ");
    const pe = bytes.readUInt32LE(0x3c);
    assert.equal(bytes.subarray(pe, pe + 4).toString("hex"), "50450000");
    assert.equal(bytes.readUInt16LE(pe + 4), arch === "x64" ? 0x8664 : 0xaa64);
  }
}

const sums = [];
for (const name of names) {
  assert.ok(expected.includes(name), `Unexpected archive: ${name}`);
  const bytes = fs.readFileSync(path.join(dir, name));
  const digest = createHash("sha256").update(bytes).digest("hex");
  const line = `${digest}  ${name}`;
  assert.equal(fs.readFileSync(path.join(dir, `${name}.sha256`), "utf8").trim(), line);
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  for (const entry of Object.values(zip.files)) {
    assert.equal(entry.unsafeOriginalName ?? entry.name, entry.name, "ZIP path was normalized");
    assert.ok(!entry.name.startsWith("/") && !entry.name.split("/").includes(".."), "Unsafe ZIP path");
  }
  const cli = name.match(/^docx-revision-analyzer-(linux|macos|windows)-(x64|arm64)\.zip$/);
  if (cli) {
    const [, os, arch] = cli;
    const binary = `docx-revision-analyzer${os === "windows" ? ".exe" : ""}`;
    assert.deepEqual(Object.keys(zip.files).sort(), ["LICENSE", binary].sort());
    checkArchitecture(await zip.file(binary).async("nodebuffer"), os, arch);
    if (os !== "windows") assert.ok((zip.file(binary).unixPermissions & 0o111) !== 0, "Executable permission missing");
  } else {
    const command = name.includes("-chart-") ? "chart" : "flow";
    const app = `docx-revision-${command}.app/`;
    assert.ok(Object.keys(zip.files).every((entry) => entry.startsWith(app)));
    assert.ok(zip.file(`${app}Contents/Resources/Scripts/main.scpt`));
    const binary = zip.file(`${app}Contents/Resources/docx-revision-analyzer`);
    assert.ok(binary);
    checkArchitecture(await binary.async("nodebuffer"), "macos", "arm64");
  }
  sums.push(line);
  console.log(`Verified ${name}`);
}
if (complete) fs.writeFileSync(path.join(dir, "SHA256SUMS.txt"), sums.join("\n") + "\n");
