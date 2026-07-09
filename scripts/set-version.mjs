import fs from "node:fs";

const version = process.argv[2];

if (!/^\d+\.\d+\.\d+$/.test(version || "")) {
  console.error("Usage: node scripts/set-version.mjs <x.y.z>");
  process.exit(1);
}

function updateJson(path, updater) {
  const data = JSON.parse(fs.readFileSync(path, "utf8"));
  updater(data);
  fs.writeFileSync(path, `${JSON.stringify(data, null, 2)}\n`);
}

function replaceFile(path, replacer) {
  fs.writeFileSync(path, replacer(fs.readFileSync(path, "utf8")));
}

updateJson("package.json", (data) => {
  data.version = version;
});

updateJson("package-lock.json", (data) => {
  data.version = version;
  if (data.packages?.[""]) {
    data.packages[""].version = version;
  }
});

updateJson("src-tauri/tauri.conf.json", (data) => {
  data.version = version;
});

replaceFile("src-tauri/Cargo.toml", (text) =>
  text.replace(/(^\[package\][\s\S]*?^version\s*=\s*)"[0-9]+\.[0-9]+\.[0-9]+"/m, `$1"${version}"`),
);

if (fs.existsSync("src-tauri/Cargo.lock")) {
  replaceFile("src-tauri/Cargo.lock", (text) =>
    text.replace(
      /(\[\[package\]\]\nname = "excel-check-tool"\nversion = )"[0-9]+\.[0-9]+\.[0-9]+"/,
      `$1"${version}"`,
    ),
  );
}
