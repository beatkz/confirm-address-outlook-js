import { resolve } from "node:path";
import { defineConfig } from "vite";
import flattenOutput from "vite-plugin-flatten-output";
import devCerts from "office-addin-dev-certs";

const devUrl = "https://localhost:3000/confirm-address-outlook-js/";
const prodUrl = "https://beatkz.github.io/confirm-address-outlook-js/";

export default defineConfig(async ({ command, mode }) => {
  const prod = mode === "production";
  const https = command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;

  return {
    base: "/confirm-address-outlook-js/",
    define: {
      "process.env.BASE_URL": JSON.stringify(prod ? prodUrl : devUrl),
    },
    plugins: [
      flattenOutput({
        removeDirs: [
          'src/bgevent', 
          'src/capopup', 
          'src/settings',
        ],
      }),
    ],
    build: {
      outDir: "dist",
      emptyOutDir: true,
      sourcemap: true,
      modulePreload: false,
      rolldownOptions: {
        input: {
          settings: resolve("src/settings/settings.html"),
          capopup: resolve("src/capopup/capopup.html"),
          bgevent: resolve("src/bgevent/bgevent.html"),
          bgevent_olc: resolve("src/bgevent/bgevent_olc.html"),
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
  };
});