/* eslint-disable no-undef */

const webpack = require('webpack'); // DefinePluginのために追加
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

//const urlDev = "https://localhost:3000/";
const urlDev = "https://beatkz.github.io/confirm-address-outlook-js/beta/";
const urlProd = "https://beatkz.github.io/confirm-address-outlook-js/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

const path = require('path');

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    entry: {
      //polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      settings: "./src/settings/settings.js",
      capopup: "./src/capopup/capopup.js",
      bgevent: "./src/bgevent/bgevent.js",
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "settings.html",
        template: "./src/settings/settings.html",
        chunks: ["settings"],
        hash: false,
      }),
      new HtmlWebpackPlugin({
        filename: "capopup.html",
        template: "./src/capopup/capopup.html",
        chunks: ["capopup"],
        hash: false,
      }),
      new HtmlWebpackPlugin({
        filename: "bgevent.html",
        template: "./src/bgevent/bgevent.html",
        chunks: [], // bgevent.js の自動挿入を防止
        hash: false,
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: path.resolve(__dirname, "assets"),
            to: path.resolve(__dirname, "dist/assets"),
          },
          {
            from: "manifest*.*",
            to: "[name][ext]",
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
      new webpack.DefinePlugin({
        'process.env.BASE_URL': JSON.stringify(dev ? urlDev : urlProd),
        'process.env.NODE_ENV': JSON.stringify(dev ? "development" : "production"),
      }),
    ],
    devtool: "source-map",
    output: {
      clean: true,
      // モードに応じてファイル名を条件付きで設定
      filename: dev ? "[name].js" : "[name].min.js",
      chunkFilename: dev ? "[name].chunk.js" : "[name].min.chunk.js", // チャンクファイル名を制御
      path: path.resolve(__dirname, "dist"),      
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
  return config;
};
