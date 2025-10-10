#!/usr/bin/env node
/**
 * ZipMail - make-dist.js
 * ----------------------
 * Creates a clean ZIP archive of the project sources using the version from package.json.
 */

import fs from "fs";
import path from "path";
import archiver from "archiver";

const packageJson = JSON.parse(fs.readFileSync("package.json", "utf8"));
const projectName = "ZipMail"; // 🔥 Always use the official name with capital Z
const version = packageJson.version || "0.0.0";
const outputFile = `${projectName}-${version}.zip`;

const output = fs.createWriteStream(outputFile);
const archive = archiver("zip", { zlib: { level: 9 } });

output.on("close", () => {
  console.log(`📦 Archive created: ${outputFile} (${archive.pointer()} bytes)`);
});

archive.on("error", (err) => {
  throw err;
});

archive.pipe(output);

// Exclude unwanted folders and files
const exclude = [
  "node_modules",
  "dist",
  ".git",
  ".vscode",
  ".DS_Store",
  "*.pem",
  "*.crt",
  "*.key",
  "npm-debug.log",
  "package-lock.json", // optional
];

const root = process.cwd();

function shouldInclude(file) {
  return !exclude.some((pattern) => file.includes(pattern));
}

function addDirectory(dir) {
  const items = fs.readdirSync(dir);
  for (const item of items) {
    const fullPath = path.join(dir, item);
    const relPath = path.relative(root, fullPath);
    const stat = fs.statSync(fullPath);
    if (!shouldInclude(relPath)) continue;
    if (stat.isDirectory()) {
      addDirectory(fullPath);
    } else {
      archive.file(fullPath, { name: relPath });
    }
  }
}

addDirectory(root);
archive.finalize();
