const path=require('path');

module.exports={
  entry:'./lib/index.ts',
  module:{
    rules:[
      {
        test:/\.ts$/,
        use:'ts-loader',
        exclude:/node_modules/,
      }
    ]
  },
  externals:{
    luxon:'luxon',
    lodash:'lodash',
    exceljs:'exceljs',
    'file-saver':'file-saver',
  },
  externalsType:'commonjs',
  resolve:{
    extensions:['.ts','.js']
  },
  output:{
    filename:'excel-tool.js',
    path:path.resolve(__dirname,'dist','bundles'),
    clean:true,
    library:{
      type:'umd',
    },
    globalObject:'this'
  },
  devtool:'source-map',
}