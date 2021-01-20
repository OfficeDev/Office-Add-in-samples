const fs = require('fs');
const path = require('path');
const os = require('os');
const webpack = require('webpack');
const webpackMerge = require('webpack-merge');
const commonConfig = require('./webpack.common.js');

// homedir() gets the user's home folder for the OS. 
// E.g., 'c:\users\[USERNAME]' for Windows or '/users/[USERNAME]' for Mac
const certPath = os.homedir() + '/.office-addin-dev-certs/';

module.exports = webpackMerge(commonConfig, {
    devtool: 'eval-source-map',
    devServer: {
        publicPath: '/',
        contentBase: path.resolve('dist'),
        hot: true,
        https: {
            key: fs.readFileSync(certPath + 'localhost.key'),
            cert: fs.readFileSync(certPath + 'localhost.crt'),
            cacert: fs.readFileSync(certPath + 'ca.crt')
        },
        compress: true,
        overlay: {
            warnings: false,
            errors: true
        },
        port: 3000,
        historyApiFallback: true
    },
    plugins: [
        new webpack.HotModuleReplacementPlugin()
    ]
});
