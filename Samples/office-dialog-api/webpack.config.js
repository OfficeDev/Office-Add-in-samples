const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = async (env = {}) => ({
  devtool: "source-map",
  entry: {
    taskpane: "./src/taskpane/taskpane.ts",
    commands: "./src/commands/commands.ts",
    dialog: "./src/dialog/dialog.ts",
  },
  output: {
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: {
          loader: "babel-loader",
          options: {
            presets: ["@babel/preset-typescript"],
          },
        },
      },
      {
        test: /\.html$/,
        use: "html-loader",
      },
      {
        test: /\.css$/,
        type: "asset/resource",
        generator: { filename: "[name][ext]" },
      },
    ],
  },
  plugins: [
    ...["taskpane", "commands", "dialog"].map(
      (page) =>
        new HtmlWebpackPlugin({
          filename: `${page}.html`,
          template: `./src/${page}/${page}.html`,
          chunks: [page],
        })
    ),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets/*.png", to: "assets/[name][ext]" },
        { from: "manifest-configurations", to: "manifest-configurations" },
        { from: "manifest.json", to: "manifest.json" },
      ],
    }),
  ],
  devServer: {
    headers: { "Access-Control-Allow-Origin": "*" },
    server: {
      type: "https",
      options: env.WEBPACK_SERVE ? await devCerts.getHttpsServerOptions() : {},
    },
    port: process.env.npm_package_config_dev_server_port || 3000,
  },
});
