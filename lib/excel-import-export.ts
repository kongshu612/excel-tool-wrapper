import {
  ICellMetaData,
  ISheetMetaData,
  IXlsxMetaData,
  IXlsxData,
  ISheetMetaDataWithRows,
} from './excel-model';

import { CellValue, Workbook, Worksheet } from 'exceljs';
import { saveAs } from 'file-saver';
import _ from 'lodash';
import { isNameEquals, getDefaultMappingByValueType, getDefaultMapping, constructSheetColumnMetasFromObject, getYesNoMapping, yesNoOptions } from './excel-helper';



// num is 0 based
function convertToXlsxColumnIndex(num: number) {
  let result = '';
  let [mod, remaining] = [0, num];
  while (remaining >= 0) {
    [mod, remaining] = [remaining % 26, Math.floor(remaining / 26) - 1];
    let c = String.fromCharCode('A'.charCodeAt(0) + mod);
    result = `${c}${result}`;
  }
  return result;
}

function findSheetMetaByName(index: number, sheetName: string, fileMetas: ISheetMetaData[]) {
  // index take high priority
  if (index <= fileMetas.length) {
    const meta = fileMetas[index - 1];
    if (isNameEquals(meta.sheetName, sheetName)) {
      return meta;
    }
  }
  // by name as second priority
  const meta = fileMetas.find(it => isNameEquals(it.sheetName, sheetName));
  return meta ?? fileMetas?.[index - 1];
}

function parseSheets(wb: Workbook, fileMetas?: ISheetMetaData[]): IXlsxData {
  let result: IXlsxData = {};
  wb.eachSheet((ws, index) => {
    // index in excel is 1 based
    let sheetMeta: ISheetMetaData | undefined;
    if (!fileMetas?.length || (sheetMeta = findSheetMetaByName(index, ws.name, fileMetas)) == null) {
      Object.assign(result, { [ws.name]: parseSheet(ws) });
      return;
    }
    const { sheetName, columns, showHeader = true, runningInStrictMode = true } = sheetMeta;
    Object.assign(result, { [sheetName]: parseSheet(ws, columns, showHeader, runningInStrictMode) });
  });
  return result;
}

function parseSheetsInCompatiableMode(wb: Workbook): IXlsxData {
  let result: IXlsxData = {};
  wb.eachSheet((ws, index) => {
    Object.assign(result, parseSheetInCompatiableMode(ws))
  });
  return result;
}

function parseSheetInCompatiableMode(ws: Worksheet) {
  let rowDatas: any[] = [];
  let maxColumn = 1;
  let { name } = ws;
  let columns: string[] = [];
  let rowReadEnded: boolean = false;
  ws.eachRow((row, index) => {
    if (index === 1) {
      let colIndex = 1;
      let header: string | null = null;
      do {
        header = row.getCell(colIndex).value as string;
        if (header?.length > 0) {
          columns.push(header.trim());
        }
        maxColumn = colIndex;
        colIndex++;
      } while (header?.length > 0);
      return;
    }
    if (rowReadEnded || !Array.isArray(row.values)) {
      return;
    }
    if (row.values[1] === BREAK_LINE) {
      rowReadEnded = true;
      return;
    }
    let rowData: any = {};
    for (let i = 1; i < maxColumn; i++) {
      let originalValue = row.values[i];
      if (typeof originalValue === 'object' && originalValue != null && 'text' in originalValue) {
        originalValue = originalValue['text'];
      }
      const key = columns[i - 1];
      const fromExcelValue = getDefaultMappingByValueType(originalValue).fromExcel(originalValue);
      // incase we have duplicate rows, the one with value will overwrite the one without value
      if (rowData[key] == null || (typeof rowData[key] === 'string' && !rowData[key].length)) {
        rowData[key] = fromExcelValue;
      }
    }
    if (isObjectNotEmpty(rowData)) {
      rowDatas.push(rowData);
    }
  });
  return { [name]: rowDatas };
}

function isObjectNotEmpty(obj: any) {
  return Object.values(obj).filter(it => !!it).length > 0;
}

function parseSheet(ws: Worksheet, columns?: ICellMetaData[], includeHeader: boolean = true, runningInStrictMode: boolean = true) {
  let result: any[] = [];
  let rowReadEnded: boolean = false;
  let columnMapping: any = {}; // this value is used in Compatiable Parse Mode
  let maxColumn = 1; // this value is used in Compatiable Parse Mode
  ws.eachRow((row, i) => {
    if (rowReadEnded || !Array.isArray(row.values)) {
      return;
    }
    if (row.values[1] === BREAK_LINE) {
      rowReadEnded = true;
      return;
    }
    // raw Mapping Scenario, we just mapping raw value from excel to memory, format and
    // transformation of the value is not included.
    // for raw Mapping, we will include headers
    // for row in excel, the index start of one, so we need to skip the zero item.
    if (!columns?.length) {
      result.push((row.values as CellValue[]).slice(1).map(it => {
        if (typeof it === 'object' && it != null && 'text' in it) {
          return it['text'];
        } else {
          return it;
        }
      }));
      return;
    }
    // index of excel start from one.
    // this is branch which will do the format and transfer, the logic is from the fieldmeta
    if (i === 1) {
      if (includeHeader == true && runningInStrictMode == true) {
        for (let colIndex = 0; colIndex < columns.length; colIndex++) {
          if (!isNameEquals(row.getCell(colIndex + 1).value as string, columns[colIndex].header as string)) {
            throw new Error(`header with name "${row.getCell(colIndex + 1).value}" is different with what from meta "${columns[colIndex].header}"`);
          }
        }
      } else if (!runningInStrictMode) {
        if (includeHeader == false) {
          throw new Error(`we don't support mapping without header in Non Strict Mode. It is easily to make a miss-mapping without header.`);
        }
        let colIndex = 1, header: string | null = null;
        do {
          header = row.getCell(colIndex).value as string;
          if (header?.length > 0) {
            const column = columns.find(it => isNameEquals(it.header!, header!));
            if (column != null) {
              columnMapping[colIndex] = column;
            } else {
              console.warn(`${header} Not Founded, we will ignore this column`);
            }
          }
          maxColumn = colIndex;
          colIndex++;
        } while (header?.length > 0);
      }
      if (includeHeader) {
        return;
      }
    }
    let rowData: any = {};
    if (runningInStrictMode) {
      columns.forEach((col, index) => {
        const { fieldName, fromExcel, dataType } = col;
        let originalValue = (row.values as CellValue[])[index + 1];
        if (typeof originalValue === 'object' && originalValue != null && 'text' in originalValue) {
          originalValue = originalValue['text'];
        }
        rowData[fieldName] = (fromExcel ?? getDefaultMapping(dataType, originalValue).fromExcel)(originalValue);
      });
    } else {
      for (let index = 1; index < maxColumn; index++) {
        const col = columnMapping[index];
        if (col != null) {
          const { fieldName, fromexcel, dataType } = col;
          let originalValue = row.values[index];
          if (typeof originalValue === 'object' && originalValue != null && 'text' in originalValue) {
            originalValue = originalValue['text'];
          }
          rowData[fieldName] = (fromexcel ?? getDefaultMapping(dataType, originalValue).fromExcel)(originalValue);
        }
      }
    }
    if (isObjectNotEmpty(rowData)) {
      result.push(rowData);
    }
  });
  return result;
}

function populateSheet(sheetMeta: ISheetMetaDataWithRows, ws: Worksheet) {
  let result: any[][] = [];
  const { columns, rows = [], showHeader = true } = sheetMeta;
  if (showHeader) {
    result.push(columns.map(it => it.header || it.fieldName));
    ws.columns = columns.map(it => ({ header: it.header || it.fieldName }));
  }
  putBreakline(ws);
  // set the cell headers and apply validations if any
  columns.forEach((col, index) => {
    if (col.required) {
      const cell = ws.getCell(1, index + 1);
      cell.font = {
        color: { argb: 'ffff0000' },
        bold: true,
      }
    }
    if (!!col.dataValidations || col.dataType === 'boolean') {
      const { options, affectedRowCount = 9999 } = col.dataValidations ?? { options: yesNoOptions };
      const optionsLabels = (options as any[]).map(it => it.label);
      const formulae = getDataValidationRange(ws, optionsLabels, index);
      for (let r = showHeader ? 2 : 1; r < affectedRowCount; r++) {
        ws.getCell(r, index + 1).dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: [formulae],
          showErrorMessage: true,
        };;
      }
    }
  });
  // populate the row data if any
  if (rows.length > 0) {
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowIndex = showHeader ? i + 2 : i + 1;
      let rowValue: any[] = [];
      columns.forEach((col, index) => {
        const { fieldName, toExcel, dataType } = col;
        let val = _.get(row, fieldName);
        val = toExcel ? toExcel(val) : getDefaultMapping(dataType, val).toExcel(val);
        rowValue.push(val);
        ws.getCell(rowIndex, index + 1).value = val;
      });
      result.push(rowValue);
    }
  }
  let columnWidth: number[] = [];
  columns.forEach((col, index) => {
    const { wch } = col;
    if (wch) {
      columnWidth[index] = wch;
    }
  })
  fitToColumn(result, ws, columnWidth);
}

const LINKROW_NUMBER: number = 100000;
const BREAK_LINE: string = '__BREAKLINE__';

// we put a break line to seperate content and data validation
function putBreakline(ws: Worksheet) {
  const cell = ws.getCell(LINKROW_NUMBER - 1, 1);
  cell.value = BREAK_LINE;
  cell.font = {
    color: { argb: 'ffffffff' },
  }
}
// column number is 0 based
function getDataValidationRange(ws: Worksheet, options: string[], column: number) {
  const columnIndex = convertToXlsxColumnIndex(column);
  for (let i = 0; i < options.length; i++) {
    let rowNumber = LINKROW_NUMBER + i;
    const cell = ws.getCell(rowNumber, column + 1);
    cell.font = {
      color: { argb: 'ffffffff' },
    }
    cell.value = options[i];
  }
  return `$${columnIndex}$${LINKROW_NUMBER}:$${columnIndex}$${LINKROW_NUMBER + options.length - 1}`;
}


function fitToColumn(arrayOfArray: any[][], ws: Worksheet, columnWidth: any[] = []) {
  // get maximum character of each column
  if (!arrayOfArray?.length) {
    return;
  }
  return arrayOfArray[0].forEach((col, i) => {
    const width = columnWidth[i];
    if (width != null) {
      ws.getColumn(i + 1).width = width;
      return;
    }
    const maxlen = Math.max(
      ...arrayOfArray
        .map((a2) => a2[i]?.toString()?.length)
        .filter((it) => it > 0)
    ) + 3;
    ws.getColumn(i + 1).width = maxlen > 15 ? maxlen : 15;
  })
}

function mergeObject(source: any[]) {
  if (!source?.length) {
    throw new Error('input parameter can not be empty');
  }
  let result = Object.assign({}, source[0]);
  for (let i = 1; i < source.length; i++) {
    for (let [key, value] of Object.entries(source[i])) {
      if (value !== null && value !== undefined) {
        result[key] = value;
      }
    };
  }
  return result;
}

function removeEmptyFieldFromData(data: any, ignoreWarning: boolean) {
  let result: any = {};
  for (let [key, value] of Object.entries(data)) {
    if (value == null) {
      if (!ignoreWarning) {
        throw new Error(`${key} in parameter is empty, you must provide a value, or set ignoreWarning as true`);
      }
      console.warn(`${key} in parameter is empty, we will skip mapping for this column`)
      continue;
    }
    result[key] = value;
  }
  return result;
}

const isBrowser = new Function("try {return this===window;}catch(e){ return false;}");

async function saveToFile(workbook: Workbook, fileName: string) {
  if (isBrowser()) {
    const buffer = await workbook.xlsx.writeBuffer();
    const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const fileExtension = '.xlsx';
    const exportFileName = fileName.indexOf('.xlsx') >= 0 ? fileName : `${fileName}${fileExtension}`;

    const blob = new Blob([buffer], { type: fileType });
    saveAs(blob, exportFileName);
  } else {
    await workbook.xlsx.writeFile(fileName);
  }
}


async function exportObjectsInSingleSheet(data: any[], fileName: string, sheetName: string = 'Sheet', showHeader: boolean = true) {
  const mergedObject = mergeObject(data);
  const { columns } = constructSheetColumnMetasFromObject(mergedObject);
  const sheet: ISheetMetaDataWithRows = { columns, sheetName, showHeader, rows: data };
  await exportToExcel({ fileName, sheets: [sheet] });
}

async function exportObjectsInMultiSheet(data: { [key in string]: any[] }, fileName: string, showHeader: boolean = true) {
  const sheets: ISheetMetaDataWithRows[] = Object.entries(data)
    .filter(([key, value]) => Array.isArray(value) && value.length > 0)
    .map(([key, value]) => {
      const mergedObject = mergeObject(value);
      const { columns } = constructSheetColumnMetasFromObject(mergedObject);
      const sheet: ISheetMetaDataWithRows = { columns, sheetName: key, showHeader, rows: value };
      return sheet;
    });
  await exportToExcel({ fileName, sheets });
}


async function exportToExcel(fileMeta: IXlsxMetaData) {
  const { fileName, sheets } = fileMeta;
  const wb = new Workbook();
  for (let sheetMeta of sheets) {
    const { sheetName } = sheetMeta;
    const ws = wb.addWorksheet(sheetName);
    populateSheet(sheetMeta, ws);
  }
  await saveToFile(wb, fileName);
}


async function parseObjectsFromExcel(file: File | string, data: any, ignoreWarning: boolean = false) {
  const consolidatedObject = removeEmptyFieldFromData(data, ignoreWarning);
  const { columns } = constructSheetColumnMetasFromObject(consolidatedObject);
  const sheet: ISheetMetaDataWithRows = { columns, sheetName: '', showHeader: true, runningInStrictMode: false };
  return await importObjects(file, [sheet]);
}

export async function exportObjects(data: any[], fileName: string, showHeader: boolean, sheetName: string): Promise<void>;
export async function exportObjects(data: { [key in string]: any[] }, fileName: string, showHeader: boolean): Promise<void>;
export async function exportObjects(fileMeta: IXlsxMetaData): Promise<void>;
export async function exportObjects(data: { [key in string]: any[] } | any[] | IXlsxMetaData, fileName?: string, showHeader: boolean = true, sheetName: string = 'Sheet'): Promise<void> {
  if (Array.isArray(data)) {
    await exportObjectsInSingleSheet(data, fileName!, sheetName, showHeader);
  } else if ('fileName' in data && 'sheets' in data && Array.isArray(data.sheets) && data.sheets.length > 0 && fileName == null) {
    await exportToExcel((data as unknown) as IXlsxMetaData);
  }
  else {
    await exportObjectsInMultiSheet(data as any, fileName!, showHeader);
  }
}

export async function importObjects(file: File | string): Promise<IXlsxData | null>;
export async function importObjects(file: File | string, fileMeta: ISheetMetaData[]): Promise<IXlsxData | null>;
export async function importObjects(file: File | string, data: { [key in string]: any }, compatiableMode: boolean): Promise<IXlsxData | null>
export async function importObjects(file: File | string, fileMeta?: ISheetMetaData[] | { [key in string]: any }, compatiableMode?: boolean): Promise<IXlsxData | null> {
  if (fileMeta != null && !Array.isArray(fileMeta) && compatiableMode !== undefined) {
    return await parseObjectsFromExcel(file, fileMeta, compatiableMode);
  }
  if (file == null) {
    return Promise.resolve(null);
  } else if (typeof file === 'string') {
    return new Promise<IXlsxData | null>(async (resolve, reject) => {
      try {
        const wb = new Workbook();
        const workbook = await wb.xlsx.readFile(file);
        resolve(fileMeta != null && Array.isArray(fileMeta) && fileMeta.length > 0 ? parseSheets(workbook, fileMeta) : parseSheetsInCompatiableMode(workbook));
      } catch (err) {
        reject(err);
      }
    });
  }
  else {
    return new Promise<IXlsxData | null>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const wb = new Workbook();
          const workbook = await wb.xlsx.load(reader.result as any);
          resolve(fileMeta != null && Array.isArray(fileMeta) && fileMeta.length > 0 ? parseSheets(workbook, fileMeta) : parseSheetsInCompatiableMode(workbook));
        } catch (err) {
          reject(err);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }
}




