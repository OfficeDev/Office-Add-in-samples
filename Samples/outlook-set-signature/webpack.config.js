/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000";
//const urlProd = "https://www.contoso.com"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION
const urlProd = "https://tabbc9b35.z22.web.core.windows.net";

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
      assignpane: ["./src/taskpane/Js/taskpane_main.js", "./src/taskpane/html/assignsignature.html"],
      editpane: ["./src/taskpane/html/editsignature.html"],
      autorun: ["./src/runtime/Js/autorunshared.js", "./src/runtime/html/autorunweb.html"],
    },
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
            options: {
              presets: ["@babel/preset-env"],
            },
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
        filename: "editsignature.html",
        template: "./src/taskpane/html/editsignature.html",
        chunks: ["polyfill", "editpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "assignsignature.html",
        template: "./src/taskpane/html/assignsignature.html",
        chunks: ["polyfill", "assignpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "autorunweb.html",
        template: "./src/runtime/html/autorunweb.html",
        chunks: ["polyfill", "autorun"],
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
          {
            from: "src/runtime/Js/autorunshared.js",
            to: "[name]" + "[ext]",
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
