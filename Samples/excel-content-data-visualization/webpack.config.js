/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://officedev.github.io/Office-Add-in-samples/Samples/excel-content-data-visualization/dist/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      home: ["./src/home/home.js", "./src/home/home.html"],
      databinding: ["./src/data-binding/data-binding.js", "./src/data-binding/data-binding.html"],
      shared: ["./src/shared/shared.js"],
    },
    output: {
      clean: true,
      // Use absolute paths for production (GitHub Pages), relative for dev
      publicPath: dev ? "/" : urlProd,
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
    plugins: [
      new HtmlWebpackPlugin({
        filename: "home.html",
        template: "./src/home/home.html",
        chunks: ["polyfill", "home", "shared"],
      }),
      new HtmlWebpackPlugin({
        filename: "data-binding.html",
        template: "./src/data-binding/data-binding.html",
        chunks: ["polyfill", "databinding", "shared"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.*",
            to: "[name]" + "[ext]",
            transform(home) {
              if (dev) {
                return home;
              } else {
                return home.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
    ],
    devServer: {
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
