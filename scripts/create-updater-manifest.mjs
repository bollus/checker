import { readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const [assetRoot, version, tag, repository] = process.argv.slice(2);
if (!assetRoot || !version || !tag || !repository) {
  throw new Error("Usage: create-updater-manifest.mjs <asset-root> <version> <tag> <owner/repo>");
}

async function walk(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const target = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...await walk(target));
    else files.push(target);
  }
  return files;
}

const files = await walk(assetRoot);
const installer = files.find((file) => /setup\.exe$/i.test(file));
if (!installer) throw new Error("Windows NSIS updater installer was not found");

const signaturePath = files.find((file) => path.resolve(file) === path.resolve(`${installer}.sig`));
if (!signaturePath) throw new Error(`Updater signature was not found for ${installer}`);

const installerName = path.basename(installer);
const manifest = {
  version,
  notes: `表格核对工具 ${tag}`,
  pub_date: new Date().toISOString(),
  platforms: {
    "windows-x86_64": {
      signature: (await readFile(signaturePath, "utf8")).trim(),
      url: `https://github.com/${repository}/releases/download/${encodeURIComponent(tag)}/${encodeURIComponent(installerName)}`,
    },
  },
};

await writeFile(path.join(assetRoot, "latest.json"), `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
