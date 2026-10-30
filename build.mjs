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

// Standalone initiative extension, published under the same Pages site.
await build({...shared,entryPoints:['initiative/src/main.tsx'],outfile:'initiative/assets/main.js',minify:true,define:{'process.env.NODE_ENV':'"production"'},jsx:'automatic',loader:{'.svg':'dataurl','.png':'dataurl'},alias:{'@mui/material':'./node_modules/@mui/material/esm','@mui/icons-material':'./node_modules/@mui/icons-material/esm'}});
await build({...shared,entryPoints:['initiative/background.js'],outfile:'initiative/assets/background.js'});
