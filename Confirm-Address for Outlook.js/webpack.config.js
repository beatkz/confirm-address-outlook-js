/* eslint-disable no-undef */

const webpack = require('webpack');
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

const path = require('path');

module.exports = async (env, options) => {
  const buildState = {
    development: function(){
      var wpconf = {
        endpoint: "https://localhost:3000/",
        minify: false,
      };
      console.dir("Building for Debug. wpconf: " + JSON.stringify(wpconf));
      return wpconf;
    },
    production: function(){
      var wpconf = {
        endpoint: "https://beatkz.github.io/confirm-address-outlook-js/",
        minify: true,
      };
      console.dir("Building for Production. wpconf: " + JSON.stringify(wpconf));
      return wpconf;
    },
  };

  const { endpoint, minify } = buildState[options.mode]();

  const config = {
    entry: {
      settings: "./src/settings/settings.ts",
      capopup: "./src/capopup/capopup.ts",
      bgevent: "./src/bgevent/bgevent.ts",
      bgevent_olc: "./src/bgevent/bgevent_olc.ts",
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
              if (minify) {
                return content;
              } else {
                var releaseendpoint = "https://beatkz.github.io/confirm-address-outlook-js/";
                return content.toString().replace(new RegExp(endpoint, "g"), releaseendpoint);
              }
            },
          },
        ],
      }),
      new webpack.DefinePlugin({
        'process.env.BASE_URL': JSON.stringify(endpoint),
      }),
    ],
    devtool: "source-map",
    output: {
      clean: true,
      filename: minify ? "[name].min.js" : "[name].js",
      chunkFilename: minify ? "[name].min.chunk.js" : "[name].chunk.js",
      path: path.resolve(__dirname, "dist"),
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.(tsx?|jsx?)$/,
          exclude: /node_modules/,
          use: "babel-loader",
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