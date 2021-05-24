const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const webpack = require("webpack");

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: "@babel/polyfill",
      runtime: "./src/runtime/Js/autorunshared.js"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader", 
            options: {
              presets: ["@babel/preset-env"]
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          use: "file-loader"
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            return content;
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "autorunweb.html",
        template: "./src/runtime/Js/autorunweb.html",
        chunks: ["polyfill", "autorunweb"]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },      
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};
