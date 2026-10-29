const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://officedev.github.io/Office-Add-in-samples/Samples/outlook-encrypt-attachments/";

module.exports = async (env = {}, options = {}) => ({
  entry: {},
  output: {
    clean: true,
  },
  plugins: [
    new CopyWebpackPlugin({
      patterns: [
        { from: "src", to: "." },
        { from: "assets/*.png", to: "assets/[name][ext]" },
        {
          from: "manifest*.*",
          to: "[name][ext]",
          transform(content) {
            return options.mode === "production"
              ? content.toString().replaceAll(urlDev, urlProd)
              : content;
          },
        },
      ],
    }),
  ],
  devServer: {
    static: false,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: {
      type: "https",
      options: env.WEBPACK_SERVE ? await devCerts.getHttpsServerOptions() : undefined,
    },
    port: process.env.npm_package_config_dev_server_port || 3000,
  },
});
