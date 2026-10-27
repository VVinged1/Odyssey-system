import { build } from "esbuild";

const shared = {
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2020",
  sourcemap: false,
  legalComments: "none",
  logLevel: "info",
};

await build({
  ...shared,
  entryPoints: ["main.js"],
  outfile: "assets/main.js",
});

await build({
  ...shared,
  entryPoints: ["background.js"],
  outfile: "assets/background.js",
});

// Schema helpers are shared with GM; always rebuild its bundle as well.
await build({ ...shared, entryPoints:["gm-extension/main.js"], outfile:"gm-extension/assets/main.js" });
await build({ ...shared, entryPoints:["room-log/main.js"], outfile:"room-log/assets/main.js" });
