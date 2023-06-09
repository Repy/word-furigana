const path = require('path');
const CopyPlugin = require('copy-webpack-plugin');

module.exports = {
    mode: "development",
    devtool: "source-map",
    entry: {
        'pane/pane.js': './src/pane/pane.ts',
    },
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: '[name]',
    },
    resolve: {
        extensions: [".ts", ".js"]
    },
    module: {
        rules: [
            // all files with a `.ts` or `.tsx` extension will be handled by `ts-loader`
            { test: /\.ts$/, loader: "ts-loader" },
        ]
    },
    plugins: [
        new CopyPlugin({
            patterns: [
                { from: './src', to: '.' },
            ],
        }),
    ],
};
