#!/usr/bin/env node
// Keeps the version the same in every file that carries it. Dataverse only
// picks up a new build of the control when the manifest version goes up, so
// these must move together.
//
//   npm run version:check          fails if the files disagree
//   npm run version:set -- 2.9.6   writes 2.9.6 to all of them

const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");

// Each entry finds the version in one file. The regex's first group is the
// text before the version, the second the version itself.
const files = [
  { file: "package.json", pattern: /^(  "version": ")([^"]+)/m },
  { file: "package-lock.json", pattern: /^(  "version": ")([^"]+)/m },
  {
    file: "package-lock.json",
    pattern: /("packages": \{\s*"": \{\s*"name": "[^"]*",\s*"version": ")([^"]+)/,
  },
  { file: "Dropzone/ControlManifest.Input.xml", pattern: /(<control [^>]*\bversion=")([^"]+)/ },
  { file: "Dropzone/version.ts", pattern: /(CONTROL_VERSION = ")([^"]+)/ },
  { file: "Solution/ModernDropzone/src/Other/Solution.xml", pattern: /(<Version>)([^<]+)/ },
];

function read(entry) {
  const text = fs.readFileSync(path.join(root, entry.file), "utf8");
  const match = entry.pattern.exec(text);
  if (!match) throw new Error(`No version found in ${entry.file}`);
  return { text, version: match[2] };
}

function check() {
  const found = files.map((entry) => ({ ...entry, version: read(entry).version }));
  const versions = new Set(found.map((f) => f.version));
  if (versions.size === 1) {
    console.log(`Version ${[...versions][0]} everywhere.`);
    return;
  }
  console.error("Versions differ:");
  for (const f of found) console.error(`  ${f.version.padEnd(12)} ${f.file}`);
  console.error("Run: npm run version:set -- <version>");
  process.exit(1);
}

function set(version) {
  if (!/^\d+\.\d+\.\d+$/.test(version)) {
    console.error(`Expected a version like 2.9.6, got "${version}".`);
    process.exit(1);
  }
  for (const entry of files) {
    const file = path.join(root, entry.file);
    const { text } = read(entry);
    fs.writeFileSync(
      file,
      text.replace(entry.pattern, (_, before) => before + version),
    );
  }
  console.log(`Set version ${version}.`);
}

const [command, arg] = process.argv.slice(2);
if (command === "set") set(arg || "");
else check();
