
const { exportObjects, importObjects } = require('../dist/bundles/excel-tool');
const path = require('path');
const _ = require('lodash');
const fs = require('fs');
const { DateTime } = require('luxon')

const { faker } = require('@faker-js/faker')

function createUser() {
  return {
    age: faker.number.int({ min: 1, max: 150 }),
    userName: faker.internet.userName(),
    sex: faker.datatype.boolean(),
    birthday: faker.date.birthdate(),
    country: countryList[faker.number.int({ min: 0, max: 4 })].value,
  };
}

const countryList = [1, 2, 3, 4, 5].map(it => ({ label: `Country ${it}`, value: `country${it}` }));

const columns = [
  { fieldName: 'userName', header: 'User Name', dataType: 'string', required: true, wch: 30 },
  { fieldName: 'age', header: 'Age', dataType: 'number', required: true, wch: 10 },
  { fieldName: 'sex', header: 'Male', dataType: 'boolean', required: true, wch: 10 },
  { fieldName: 'birthday', header: 'Birthday', dataType: 'dateTime', required: true, wch: 30 },
  { fieldName: 'country', header: 'From Country', dataType: 'string', required: true, wch: 30, dataValidations: { options: countryList } },
];

const rows = faker.helpers.multiple(createUser, { count: 2 });

function getSheet(fileName) {
  return {
    fileName,
    sheets: [{ columns, sheetName: 'Sheet', showHeader: true, rows }]
  }
}

async function test1() {
  const fileName = path.join(__dirname, 'output_meta1.xlsx');
  const metas = getSheet(fileName);
  await exportObjects(metas);
  const result = await importObjects(fileName, metas.sheets);
  if (_.isEqualWith(rows, result['Sheet'], (value, other, propty) => {
    if (propty == 'birthday') {
      return DateTime.fromJSDate(value).toFormat('yyyy-MM-dd') === (other);
    }
  })) {
    console.log('######test pass for test case1#######');
  } else {

    console.error('######error for test case1######')
    console.log('original rows are:');
    console.log(rows);
    console.log('from excel are');
    console.log(result);
  }
  fs.unlinkSync(fileName);
}


async function timeZoneTesting() {
  const fileName = path.join(__dirname, `timezone.xlsx`);
  const columns = [
    { fieldName: 'birthday', header: 'Birthday', dataType: 'dateTime', required: true, wch: 30 },
  ];
  const result = await importObjects(fileName, [{ columns, sheetName: 'Sheet', showHeader: true }]);
  if (_.isEqual([{ birthday: '2024-01-07' }], result['Sheet'])) {
    console.log('######test pass for test case1#######');
  } else {
    console.error('######error for test case1######')
    console.log('original rows are:');
    console.log({ birthday: '2024-01-07' });
    console.log('from excel are');
    console.log(result);
  }
}

(async () => {
  await test1();
  await timeZoneTesting();
})();