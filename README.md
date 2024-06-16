# Excel-Tool
this packet is a js wrapper around **[exceljs](https://github.com/exceljs/exceljs)**. we provide an easy to use apis to help you import/export data to excel files without considering the details of exceljs. 

## Features
- import/export between js objects and excel file
- multiple data types support e.g number/boolean/datetime/array
- customization while exporting to excel or importing from excel
- support running both in browser and node

## PeerDependency
inorder to use this package, make sure you install these packages.
`exceljs file-saver lodash luxon`

## Usage
- install package

after setting, you can run this command to install package into your project
```
npm install excel-tool-wrapper
```

- Examples

we provide examples folders which hold examples to use this package. you can refer to the examples through the [source code](https://github.com/kongshu612/excel-tool-wrapper), Or you can access to the examples from the npm packages. Here we give some basic examples 

  + export an array into excel

incase we have a objects like this, and we hope to export this into a single sheet of excel file

```ts
import {exportObjects} from 'excel-tool-wrapper';
const rows = [
  {"header 1":"1x1","header 2":"1x2","header 3":"1x3"},
  {"header 1":"2x1","header 2":"2x2","header 3":"2x3"},
  {"header 1":"3x1","header 2":"3x2","header 3":"3x3"}
];
exportObjects(rows,'output.xlsx',true,'SheetName');
```
  + import object from excel

we will import the exported xlsx from the first example
```ts
import {importObjects} from 'excel-tool-wrapper';
const result = await importObjects(file);
result;
// result will have this structure
// {
//   "SheetName":[
//     {"header 1":"1x1","header 2":"1x2","header 3":"1x3"},
//     {"header 1":"2x1","header 2":"2x2","header 3":"2x3"},
//     {"header 1":"3x1","header 2":"3x2","header 3":"3x3"}
//   ]
// };
```

happy Coding


