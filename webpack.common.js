const path = require('path')

module.exports = {
  // devtool: 'none', // no eval()
  entry: {
    main: './src/js/main.js',
    vendor: [
      './src/js/vendor/jquery.min.js',
      './src/js/vendor/popper.min.js',
      './src/js/vendor/bootstrap.min.js',
      './src/js/vendor/vendor.js'
    ] 
  },
  module: {
    rules: [{
        test: /\.html$/,
        use: ['html-loader']
      },
      {
        test: /\.(jpe?g|png|gif|svg)$/,
        use: {
          loader: 'file-loader',
          options: {
            name: '[name].[ext]',
            outputPath: 'img'
          }
        }
      },
      {
        test: /jquery.+\.js$/,
        use: [{
            loader: 'expose-loader',
            options: 'jQuery'
        },{
            loader: 'expose-loader',
            options: '$'
        }] 
      }
    ]
  },
  externals: {
    jquery: 'jQuery'
  }
}