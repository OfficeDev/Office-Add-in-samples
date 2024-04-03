const fs = require('fs');
const path = require('path');
const os = require('os');
const webpack = require('webpack');
const { merge } = require('webpack-merge');
const commonConfig = require('./webpack.common.js');

// homedir() gets the user's home folder for the OS. 
// E.g., 'c:\users\[USERNAME]' for Windows or '/users/[USERNAME]' for Mac
const certPath = os.homedir() + '/.office-addin-dev-certs/';

module.exports = merge(commonConfig, {
    devtool: 'eval-source-map',
    devServer: {
        client: {
            overlay: {
                warnings: false,
                errors: true
            }
        },
        static: {
            directory: path.resolve('dist'),
            publicPath: '/'
        },
        server: {
            type: 'https',
            options: {
                cert: certPath + 'localhost.crt',
                cacert: certPath + 'ca.crt',
                key: certPath + 'localhost.key'
            }
        },
        compress: true,
        port: 3000,
        historyApiFallback: true
    },
    plugins: [
        new webpack.HotModuleReplacementPlugin()
    ]
});
