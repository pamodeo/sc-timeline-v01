const path = require('path');

module.exports = {
  entry: {
    taskpane: './taskpane.js',
    commands: './commands.js'
  },
  output: {
    filename: '[name].js',
    path: path.resolve(__dirname, 'dist')
  },
  devServer: {
    static: {
      directory: path.join(__dirname, './')
    },
    port: 3000,
    server: 'https'
  },
  mode: 'development'
};