const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const webpackMerge = require('webpack-merge');
const commonConfig = require('./webpack.common.js');

const defaultCertDirectory = require('os').homedir();
const certFilePath = path.resolve(defaultCertDirectory, '.office-addin-dev-certs', 'localhost.crt');
const keyFilePath = path.resolve(defaultCertDirectory, '.office-addin-dev-certs', 'localhost.key');
const caFilePath = path.resolve(defaultCertDirectory, '.office-addin-dev-certs', 'ca.crt');

module.exports = webpackMerge(commonConfig, {
    devtool: 'eval-source-map',
    devServer: {
        headers: {
            "Access-Control-Allow-Origin": "*"
          },  
        publicPath: '/',
        contentBase: path.resolve('dist'),
        hot: true,
        https: {
            key: fs.readFileSync(keyFilePath),
            cert: fs.readFileSync(certFilePath),
            cacert: fs.readFileSync(caFilePath)
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
