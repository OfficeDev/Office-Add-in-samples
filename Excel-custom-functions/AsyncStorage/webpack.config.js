// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const path = require('path');

module.exports = {
    mode: 'development',
    entry: './src/functions/functions.js',
    output: {
        path: path.resolve(__dirname, 'dist/win32/ship'),
        filename: 'index.win32.bundle'
    },
    devtool: "source-map",
    resolve: {
        extensions: ['.js', 'json']
    },
    devServer: {
        port: 8081,
        hot: true,
        inline: true,
        headers: {
            "Access-Control-Allow-Origin": "*"
        }
    }
};