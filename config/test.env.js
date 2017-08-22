var merge = require('webpack-merge')
var devEnv = require('./dev.env.js')

module.exports = merge(devEnv, {
  NODE_ENV: '"testing"'
})
