const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require("webpack");

const CleanWebpackPlugin = require("clean-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");


module.exports = {
    devtool: 'source-map',
    entry: {
        app: './src/index.ts',
        functions: "./src/functions/functions.ts",
        polyfill: "@babel/polyfill",
    },
    resolve: {
        extensions: ['.ts', '.tsx', '.html', '.js']
    },
    module: {
        rules: [
            {
                test: /\.ts$/,
                exclude: /node_modules/,
                use: "babel-loader"
              },
            {
                test: /\.tsx?$/,
                exclude: /node_modules/,
                use: 'ts-loader'
            },
            {
                test: /\.html$/,
                exclude: /node_modules/,
                use: 'html-loader'
            },
            {
                test: /\.(png|jpg|jpeg|gif)$/,
                use: 'file-loader'
            }
        ]
    },
    plugins: [
        new CleanWebpackPlugin({
            cleanOnceBeforeBuildPatterns: dev ? [] : ["**/*"]
          }),
          new CustomFunctionsMetadataPlugin({
            output: "functions.json",
            input: "./src/functions/functions.ts"
          }),
          new HtmlWebpackPlugin({
            filename: "functions.html",
            template: "./src/functions/functions.html",
            chunks: ["polyfill", "functions"]
          }),
        new HtmlWebpackPlugin({
            template: './index.html',
            chunks: ['app']
        }),
        new HtmlWebpackPlugin({
            template: './logout/logout.html',
            filename: 'logout/logout.html',
            chunks: ['logout']
        }),
        new HtmlWebpackPlugin({
            template: './logoutcomplete/logoutcomplete.html',
            filename: 'logoutcomplete/logoutcomplete.html',
            chunks: ['logoutcomplete']
        }),
        new webpack.ProvidePlugin({
            Promise: ["es6-promise", "Promise"]
        })
    ]
};