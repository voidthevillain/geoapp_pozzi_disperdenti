const path = require('path')
const common = require('./webpack.common')
const merge = require('webpack-merge')
const HtmlWebpackPlugin = require('html-webpack-plugin')

module.exports = merge(common, {
  mode: 'development',
  // devtool: 'none', // no eval()
  output: {
    filename: '[name].bundle.js',
    path: path.resolve(__dirname, 'dist')
  },
  plugins: [
    new HtmlWebpackPlugin({
    template: './src/index.html'
  })],
  module: {
    rules: [
      {
        test: /\.css$/,
        use: [
          'style-loader', // Inject styles into DOM
          'css-loader'    // Turns css into commonjs
        ]
      },
    ]
  }
})