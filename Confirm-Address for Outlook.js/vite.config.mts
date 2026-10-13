import { defineConfig } from "vite";
import path from "node:path";
import { fileURLToPath } from "node:url";
import devCerts from "office-addin-dev-certs";

const root = path.dirname(fileURLToPath(import.meta.url));
const devUrl = "https://localhost:3000/";
const prodUrl = "https://beatkz.github.io/confirm-address-outlook-js/";

const htmlMap: Record<string, string> = {
  "/settings.html": "/src/settings/settings.html",
  "/capopup.html": "/src/capopup/capopup.html",
  "/bgevent.html": "/src/bgevent/bgevent.html",
  "/bgevent_olc.html": "/src/bgevent/bgevent_olc.html",
};

export default defineConfig(async ({ command, mode }) => {
  const prod = mode === "production";
  const https = command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;

  return {
    base: prod ? "/confirm-address-outlook-js/" : "/",
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
      https: https ? { ca: https.ca, key: https.key, cert: https.cert } : undefined,
      headers: { "Access-Control-Allow-Origin": "*" },
    },
    plugins: [
      {
        name: "office-flat-paths",
        configureServer(server) {
          server.middlewares.use(async (req, _res, next) => {
            const url = req.url?.split("?")[0] ?? "";
            if (htmlMap[url]) req.url = htmlMap[url];
            next();
          });
        },
        generateBundle(_options, bundle) {
          for (const file of Object.keys(bundle)) {
            if (!file.endsWith(".html") || !file.includes("/")) continue;
            const flat = file.split("/").pop()!;
            bundle[file].fileName = flat; // 出力ファイル名自体を更新
            bundle[flat] = bundle[file];
            delete bundle[file];
          }
        },
      },
    ],
  };
});