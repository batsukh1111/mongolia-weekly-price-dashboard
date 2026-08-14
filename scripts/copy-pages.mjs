import { cpSync, mkdirSync, readdirSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const root = process.cwd();
const dist = join(root, "dist");
const docs = join(root, "docs");
const assets = join(root, "assets");

rmSync(docs, { recursive: true, force: true });
rmSync(assets, { recursive: true, force: true });
mkdirSync(docs, { recursive: true });

for (const name of readdirSync(dist)) {
  cpSync(join(dist, name), join(docs, name), { recursive: true });
}

const builtHtml = join(docs, "index.src.html");
const docsHtml = join(docs, "index.html");
try {
  renameSync(builtHtml, docsHtml);
} catch {
  // already named index.html
}

cpSync(docsHtml, join(root, "index.html"));
if (readdirSync(docs).includes("assets")) {
  cpSync(join(docs, "assets"), assets, { recursive: true });
}
if (readdirSync(docs).includes("favicon.svg")) {
  cpSync(join(docs, "favicon.svg"), join(root, "favicon.svg"));
}
writeFileSync(join(root, ".nojekyll"), "");
writeFileSync(join(docs, ".nojekyll"), "");
console.log("Copied production files to / and docs/");
