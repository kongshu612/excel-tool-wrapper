const basicRows=[1,2,3,4,5,6,7,8,9,10].map(row=>{
  return [1,2,3,4,5,].reduce((pre,col)=>{
    return {
      ...pre,
      [`Header ${col}`]:`${row}x${col}`,
    }
  },{});
});

function getSheet(num){
  return [1,2,3,4,5,6,7,8,9,10].map(row=>{
    return [1,2,3,4,5,].reduce((pre,col)=>{
      return {
        ...pre,
        [`Header ${col}`]:`${num}x${row}x${col}`,
      }
    },{});
  })
}

const rowsWithZip=[
  {'header 1':'1x1','header 3':'1x3','header 5':'1x5'},
  {'header 2':'2x2','header 4':'2x4','header 5':'2x5'}
];

const twoSheets={
  'Sheet1':getSheet(1),
  'Sheet2':getSheet(2),
}

const twoSheetsWithZip={
  'Sheet1':rowsWithZip,
  'Sheet2':rowsWithZip,
}


module.exports={
  basicRows,
  rowsWithZip,
  twoSheets,
  twoSheetsWithZip,
}