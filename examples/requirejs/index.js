requirejs.config({
  paths:{
    lodash:'https://cdn.jsdelivr.net/npm/lodash@4.17.21/lodash.min',
    luxon:'https://cdn.jsdelivr.net/npm/luxon@3.3.0/build/global/luxon.min',
    'file-saver':'./filesaver',
    exceljs:'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min',
    'excel-tool':'../../dist/bundles/excel-tool'
  }
});



requirejs(['excel-tool','data-source'],function(exceltool,fake){
  
  const btn =document.getElementById('export');
  const exportfunc=async ()=>{
    const fileName='ouput.xlsx';
    await exceltool.exportObjects(fake.basicRows,fileName,true,'Sheet1');
  };
  btn.addEventListener('click',exportfunc);
})