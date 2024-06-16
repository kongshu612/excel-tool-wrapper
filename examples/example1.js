

const {exportObjects,importObjects}=require('../dist/bundles/excel-tool');
const path=require('path');
const _ = require('lodash');
const fs = require('fs');
const {basicRows:rows,rowsWithZip,twoSheets,twoSheetsWithZip}=require('./fack')

async function testcase1(){
  const filename = path.join(__dirname,'output1.xlsx');
  await exportObjects(rows,filename,true,'Sheet1');
  const result = await importObjects(filename);
  if(compareArray(rows,result['Sheet1'])){
    console.log('######test pass for test case1#######');
  }else{
    
    console.error('######error for test case1######')
    console.log('original rows are:');
    console.log(rows);
    console.log('from excel are');
    console.log(result);
  }
  fs.unlinkSync(filename);
}

function compareArray(obj,other){
  if(obj.length!=other.length){
    return false;
  }
  for(let i=0;i<obj.length;i++){
    const [a,b]=[obj[i],other[i]];
    for(let key of Object.keys(b)){
      const [s,t]=[a[key],b[key]];
      if([null,'',undefined].includes(s)&&[null,'',undefined].includes(t)){
        continue;
      }else if(s!=t){
        return false;
      }
    }
  }
  return true;
}

function compareTwoSheet(obj,other){
  for(let key of Object.keys(other)){
    if(!compareArray(obj[key],other[key])){
      return false;
    }
  }
  return true;
}

async function testcase2(){
  const filename = path.join(__dirname,'output2.xlsx');
  await exportObjects(rowsWithZip,filename,true,'Sheet1');
  const result = await importObjects(filename);
  if(compareArray(rowsWithZip,result['Sheet1'])){
    console.log('######test pass for test case 2#######');
  }else{
    console.error('######error for test case2######')
    console.log('original rows are:');
    console.log(rowsWithZip);
    console.log('from excel are');
    console.log(result);
  }
  fs.unlinkSync(filename);
}


async function testcase3(){
  const filename = path.join(__dirname,'output3.xlsx');
  await exportObjects(twoSheets,filename,true);
  const result = await importObjects(filename);
  if(_.isEqual(twoSheets,result)){
    console.log('######test pass for test case3#######');
  }else{
    
    console.error('######error for test case3######')
    console.log('original rows are:');
    console.log(twoSheets);
    console.log('from excel are');
    console.log(result);
  }
  fs.unlinkSync(filename);
}


async function testcase4(){
  const filename = path.join(__dirname,'output4.xlsx');
  await exportObjects(twoSheetsWithZip,filename,true);
  const result = await importObjects(filename);
  if(compareTwoSheet(twoSheetsWithZip,result)){
    console.log('######test pass for test case4#######');
  }else{
    
    console.error('######error for test case4######')
    console.log('original rows are:');
    console.log(twoSheetsWithZip);
    console.log('from excel are');
    console.log(result);
  }
  fs.unlinkSync(filename);
}



(async()=>{
  await testcase1();
  await testcase2();
  await testcase3();
  await testcase4();
})()


