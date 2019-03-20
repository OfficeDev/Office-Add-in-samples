const path = require('path');

module.exports = {
    mode: 'development',
    entry: './src/customfunctions.js',
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