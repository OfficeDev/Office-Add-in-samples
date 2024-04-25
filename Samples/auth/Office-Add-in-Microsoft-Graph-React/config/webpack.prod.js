const webpack = require('webpack');
const { merge } = require('webpack-merge');
const commonConfig = require('./webpack.common.js');
const ENV = process.env.NODE_ENV = process.env.ENV = 'production';

module.exports = merge(commonConfig, {
    devtool: 'source-map',

    performance: {
        hints: "warning"
    },

    optimization: {
        minimize: true
    }
});

