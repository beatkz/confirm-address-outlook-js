import { defineConfig } from "vite";
import path from "node:path";
import { fileURLToPath } from "node:url";
import devCerts from "office-addin-dev-certs";
import esbuild from "esbuild";

const root = path.dirname(fileURLToPath(import.meta.url));
const devUrl = "https://localhost:3000/";
const prodUrl = "https://beatkz.github.io/confirm-address-outlook-js/";

const htmlMap: Record<string, string> = {
  "/settings.html": "/src/settings/settings.html",
  "/capopup.html": "/src/capopup/capopup.html",
  "/bgevent.html": "/src/bgevent/bgevent.html",
  "/bgevent_olc.html": "/src/bgevent/bgevent_olc.html",
};

export default defineConfig(async ({ mode }) => {
  const prod = mode === "production";
  const https = await devCerts.getHttpsServerOptions();

  return {
    define: {
      "process.env.BASE_URL": JSON.stringify(prod ? prodUrl : devUrl),
    },
    build: {
      outDir: "dist",
      emptyOutDir: true,
      sourcemap: true,
      modulePreload: false,
      rollupOptions: {
        input: {
          settings: path.resolve(root, "src/settings/settings.html"),
          capopup: path.resolve(root, "src/capopup/capopup.html"),
          bgevent: path.resolve(root, "src/bgevent/bgevent.html"),
          bgevent_olc: path.resolve(root, "src/bgevent/bgevent_olc.html"),
        },
        output: {
          entryFileNames: prod ? "[name].min.js" : "[name].js",
          chunkFileNames: prod ? "[name].min.chunk.js" : "[name].chunk.js",
          assetFileNames: "assets/[name][extname]",
        },
      },
    },
    server: {
      port: 3000,
      strictPort: true,
      https: { ca: https.ca, key: https.key, cert: https.cert },
      headers: { "Access-Control-Allow-Origin": "*" },
    },
    plugins: [
      {
        name: "office-flat-paths",
        configureServer(server) {
          server.middlewares.use(async (req, _res, next) => {
            const url = req.url?.split("?")[0] ?? "";
            if (htmlMap[url]) req.url = htmlMap[url];
            if (url === "/bgevent_olc.js") {
              const result = await esbuild.build({
                absWorkingDir: root,
                entryPoints: ["src/bgevent/bgevent_olc.ts"],
                bundle: true,
                format: "iife",
                platform: "browser",
                write: false,
                sourcemap: "inline",
              });
              _res.setHeader("Content-Type", "text/javascript");
              _res.end(result.outputFiles[0].text);
              return;
            }
            next();
          });
        },
        generateBundle(_options, bundle) {
          for (const file of Object.keys(bundle)) {
            if (!file.endsWith(".html") || !file.includes("/")) continue;
            const flat = file.split("/").pop()!;
            bundle[flat] = bundle[file];
            delete bundle[file];
          }
        },
      },
    ],
  };
});