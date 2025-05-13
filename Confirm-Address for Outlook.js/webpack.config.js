/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://github.com/beatkz/confirm-address-outlook-js/"; 

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const config = {
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      properties: ["./src/properties/properties.js", "./src/properties/properties.html"],
      settings: "./src/settings/settings.js",
      confirm: "./src/confirm/confirm.js",
      bgevent: "./src/bgevent/bgevent.js",
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "properties.html",
        template: "./src/properties/properties.html",
        chunks: ["polyfill", "properties"],
      }),
      new HtmlWebpackPlugin({
        filename: "settings.html",
        template: "./src/settings/settings.html",
        chunks: ["settings"],
      }),     new HtmlWebpackPlugin({
        filename: "confirm.html",
        template: "./src/confirm/confirm.html",
        chunks: ["confirm"],
      }),
      new HtmlWebpackPlugin({
        filename: "bgevent.html",
        template: "./src/bgevent/bgevent.html",
        chunks: ["bgevent"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
    ],
    devtool: "source-map",
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    devServer: {
      hot: false,
      liveReload: false,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
  const dev = options.mode === "development";
  return config;
};
