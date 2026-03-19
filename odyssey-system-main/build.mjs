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
