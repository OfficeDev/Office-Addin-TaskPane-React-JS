/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const path = require("path");
const webpack = require("webpack");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  // const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      vendor: ["react", "react-dom", "core-js", "@fluentui/react"],
      taskpane: ["react-hot-loader/patch", path.resolve(__dirname, "./src/test.index.tsx")],
    },
    output: {
      path: path.resolve(__dirname, "testBuild"),
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      fallback: {
        child_process: false,
        fs: false,
        os: require.resolve("os-browserify/browser"),
      },
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: [
            "react-hot-loader/webpack",
            {
              loader: "babel-loader",
              options: {
                presets: ["@babel/preset-env"],
              },
            },
          ],
          exclude: /node_modules/,
        },
        {
          test: /\.tsx?$/,
          use: ["react-hot-loader/webpack", "ts-loader"],
          exclude: /node_modules/,
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
    plugins: [
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
        process: "process/browser",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: path.resolve(__dirname, "./src/test-taskpane.html"),
        chunks: ["taskpane", "vendor", "polyfill"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/icon-*",
            to: "assets/[name][ext][query]",
          },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: "testBuild",
      },
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
