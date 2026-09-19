import { build } from "esbuild";

await build({
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2020",
  sourcemap: false,
  legalComments: "none",
  logLevel: "info",
  entryPoints: ["main.js"],
  outfile: "assets/main.js",
});
