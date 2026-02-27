/* eslint-disable no-undef */

const webpack = require('webpack');
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlStable = "https://beatkz.github.io/confirm-address-outlook-js/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

const path = require('path');

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    entry: {
      settings: "./src/settings/settings.js",
      capopup: "./src/capopup/capopup.js",
      bgevent: "./src/bgevent/bgevent.js",
      bgevent_olc: "./src/bgevent/bgevent_olc.js", // ★ 追加：Outlook Classic用エントリーポイント
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
        chunks: ["bgevent"],
        hash: false,
      }),
      // ★ Outlook Classic用HTMLファイル（bgevent_olc.html）
      // → bgevent_olc.htmlはダミーファイルで、実際のイベントハンドラーはbgevent_olc.jsに記述する想定
      new HtmlWebpackPlugin({
        filename: "bgevent_olc.html",
        template: "./src/bgevent/bgevent.html",
        chunks: ["bgevent_olc"],
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
                return content.toString().replace(new RegExp(urlDev, "g"), urlStable);
              }
            },
          },
        ],
      }),
      new webpack.DefinePlugin({
        'process.env.BASE_URL': JSON.stringify(dev ? urlDev : urlStable),
      }),
    ],
    devtool: "source-map",
    output: {
      clean: true,
      filename: dev ? "[name].js" : "[name].min.js",
      chunkFilename: dev ? "[name].chunk.js" : "[name].min.chunk.js",
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
      liveReload: true,
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