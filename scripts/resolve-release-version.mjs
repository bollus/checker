import fs from "node:fs";

function parseVersion(value) {
  const match = String(value || "").trim().match(/^v?(\d+)\.(\d+)\.(\d+)$/);
  return match ? match.slice(1).map(Number) : null;
}

function compareVersions(left, right) {
  for (let index = 0; index < 3; index += 1) {
    if (left[index] !== right[index]) {
      return left[index] - right[index];
    }
  }
  return 0;
}

function formatVersion(version) {
  return version.join(".");
}

const manifestVersion = parseVersion(JSON.parse(fs.readFileSync("package.json", "utf8")).version);
if (!manifestVersion) {
  throw new Error("package.json version must be x.y.z");
}

const latestTagVersion = parseVersion(process.argv[2]);
const nextTagVersion = latestTagVersion
  ? [latestTagVersion[0], latestTagVersion[1], latestTagVersion[2] + 1]
  : manifestVersion;
const releaseVersion = compareVersions(manifestVersion, nextTagVersion) >= 0 ? manifestVersion : nextTagVersion;
const releaseVersionText = formatVersion(releaseVersion);
const output = `version=${releaseVersionText}\ntag=v${releaseVersionText}\n`;

if (process.env.GITHUB_OUTPUT) {
  fs.appendFileSync(process.env.GITHUB_OUTPUT, output);
} else {
  process.stdout.write(output);
}
