# Motivation
excel is a powerfull tool which you can't bypass when you are doing business for financial. we are always be faced with requirement like exporting some data into excel file, then business guys make modifications manually in excel and at last they upload it into system. Excel is powerful and flexiable, so integrating excel into our applicaiton will make our applications more powerful. Here, I build this library. For our front end developers, we can do the exporting and parsing of excel without the help from backend API. It will be more cool and faster for our application. Based on the **[exceljs](https://github.com/exceljs/exceljs)**, I build this library. this library abstract the API call from excel. you don't need to care about the excel knowledge and you even don't need to care about the APIs of excel. what you need to do is just giving us the data you want to export and then library will handle it for you. And also, if the excel file is provided and the we can give you the data. That's it.

# Assumption
In order to make things easy, we give this assumption. for each sheet of excel, we will give you an array. each row of a single sheet will be parse as an single object. for the first row of the excel, we will treat it as column names. and each column name will be treated as property name of the object. If we have multiple sheets, we will give you an object like this format `{[sheetName]:Array<T>}`.

# Features
- import/export between js objects and excel file
- intutive style and type safe guarantee
- multiple data types support out of box and easy to extend e.g number/boolean/datetime/array
- data format support in excel e.g color/font-bold/cell options
- default transfer logic can cover most cases and easy to extend on the property level.
- customization while exporting to excel or importing from excel
- smart enough to understand your excel and intutive parse logic
- multiple platform support (both node and browser)

# PeerDependency
this package has peerdependency on these packages
`exceljs file-saver lodash luxon`. you don't need to install them manually, install this package will auto include them into your project.

# Usage
- install package

you can run this command to install package into your project
```
npm install excel-tool-wrapper
```

## Example and API
we provide examples folders which hold examples to use this package. you can refer to the examples through the **[Example](https://github.com/kongshu612/excel-tool-wrapper/tree/main/examples)**, Or you can access to the examples from the npm packages. Here we give some basic examples
you can refer to the  for more details

### Export an array into excel

in case we have a objects like this, and we hope to export this into a single sheet of excel file. just a single function call can solve all the things.

```ts
import {exportObjects} from 'excel-tool-wrapper';
const rows = [
  {"header 1":"1x1","header 2":"1x2","header 3":"1x3"},
  {"header 1":"2x1","header 2":"2x2","header 3":"2x3"},
  {"header 1":"3x1","header 2":"3x2","header 3":"3x3"}
];
exportObjects(rows,'output.xlsx',true,'SheetName');
```
### Params
- `async function exportObjects(data: any[], fileName: string, showHeader: boolean, sheetName: string): Promise<void>`

this is the simplest usage of export, just provide the data and the excel file name and sheetname. this library will analysis the **data** you provided automatically and every enumerable properties will be used as the column header.

**Note:** *if the data you provide don't have the same schema, library will introduce the all sets of the properties they have. for object which don't have the extra properties, will be exported as blank.* 

- `async function exportObjects(data: { [key in string]: any[] }, fileName: string, showHeader: boolean): Promise<void>;`

we can support multiple sheets exporting. just provide the data, the library will handle it automatically.

- `async function exportObjects(fileMeta: IXlsxMetaData): Promise<void>`

we still provide the more complex APIs, you can provide the extension from property level. for every property, if you don't provide the customization logic, we will use default logic. but your cusomization will take high priority.

- here is the detail definition of the `IXlxsMetaData`
```ts
export interface IXlsxMetaData {
  sheets: ISheetMetaDataWithRows[];
  fileName: string;
}
export interface ISheetMetaData {
  columns: ICellMetaData[];
  sheetName: string;
  showHeader?: boolean;
  runningInStrictMode?: boolean;
}
export interface ISheetMetaDataWithRows extends ISheetMetaData {
  rows?: any[];
}

```
-  Description of `ISheetMetaDataWithRows`

|Parameter|Description|Default Value|
|--|--|--|
|columns|Full description for each Cell|null|
|sheetName|the name of sheet|by default it is the propertyName from the root level of objects|
|showHeader|flag to determine if display column name in the first page|true|
|runningInStrictMode|internal flag, it is used while paring excel, it will be used to determine if conflict found between data and meta data,how can we process it. true, meta data will take high priority, false, data will take high priority|false|
|rows|the data to be exported into excel||

- here is the detail definition of `ICellMetaData`
```ts
export interface ICellMetaData extends Partial<IXlsxTranfer> {
  fieldName: string;
  dataValidations?: ICellDataValidation;
  required?: boolean;
  dataType?: CellDataType;
  header?: string;
  wch?: number;
}

export interface IXlsxTranfer {
  toExcel: (val: any) => any;
  fromExcel: (val: any) => any;
}
```

|Parameter|Description|Default Value|
|--|--|--|
|fieldName|the field name of this column, we will read the value of this field to feed the cell||
|dataValidations|data validations, it is the concept of excel. it define the legal options of this cell. you can also specify the row numbers will follow this rule|null|
|required|specify if the cell is mandatory, true, we will apply a mandatory style to the text.e.g. make this column name color red. and also, we will append the logic while parsing the data|false|
|dataType|specify the data type of this cell. e.g. string,number,date,boolean. we will do meaningfull transfer to keep the intuitive style and type safe from js level||
|header|the column name of the excel, by default, we will use the fieldName. but, you can a more meaningfull one||
|wch|this will be used to specify the width of the cell. In most cases, we don't need this one, As we will do the calcuation automatically. but, we do see some edge cases like, one cell has a very long text. so we five this parameter to avoid such edge case.||
|toExcel|we notice that, in some edge case, we need more complex logic while handling the transfer, so here we provide an interface to let you to customize the transfer logic. you can provide a hook function and we will call it at the last step to write your data into excel||
|fromExcel|similar to toExcel, we still give you a chance to intercept the parsing logic while parsing the excel||

### Import object from excel

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

this is the simplest cases for parsing excel file. In most cases, you only need this one. our library will handle all the transfer logic for you automatically. but we provide more complex API

- `async function importObjects(file: File | string, fileMeta: ISheetMetaData[]): Promise<IXlsxData | null>;`

like the *export API*, you can also provide the meta data to help us to understand the excel file so that we can give the data you want. the meta data file format is the same as the *export API*

- `async function importObjects(file: File | string, data: { [key in string]: any }, compatiableMode: boolean): Promise<IXlsxData | null>`

this API is an advance API. in stead of the metadata, you can just give us an example of the data you want to get. or you can give us a default data. library will use the default data to collect and analysis the data you want to get from the excel file. sometimes, we notice that proivde a metadata list is a big and tedious job. but we can easy make such a default data.

## Support this package
If you like this package, consider giving it a github star ⭐

Also, PR's are welcome!


