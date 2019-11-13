const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require("webpack");

module.exports = {
    devtool: 'source-map',
    entry: {
        app: './src/index.ts',
        'function-file': './function-file/function-file.ts',


        'login': './login/login.ts',
        'logout': './logout/logout.ts',
    },
    resolve: {
        extensions: ['.ts', '.tsx', '.html', '.js']
    },
    module: {
        rules: [
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
        new HtmlWebpackPlugin({
            template: './index.html',
            chunks: ['app']
        }),
        new HtmlWebpackPlugin({
            template: './function-file/function-file.html',
            filename: 'function-file/function-file.html',
            chunks: ['function-file']
        }),
        new HtmlWebpackPlugin({
            template: './login/login.html',
            filename: 'login/login.html',
            chunks: ['login']
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