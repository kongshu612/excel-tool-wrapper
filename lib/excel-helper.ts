import { CellDataType, ICellMetaData, ITransferWithDataType, ISheetMetaDataWithRows } from "./excel-model";
import { DateTime } from "luxon";


let fakeId: number = -1;

export const yesNoOptions = [{ label: 'Yes', value: true }, { label: 'No', value: false }];

function getFakeId() {
  return fakeId--;
}

export function isNameEquals(original: string, other: string) {
  return original?.toLocaleLowerCase?.()?.trim() == other?.toLocaleLowerCase?.()?.trim();
}

export function getLookupCodeMapping(
  options: { label: string; value: any }[],
  toExcelDefault = null,
  dataType: CellDataType = 'number',
): ITransferWithDataType {
  return <ITransferWithDataType>{
    toExcel: (val?: any) => options.find((it) => it.value === val)?.label ?? toExcelDefault ?? (typeof (val) === 'string' ? val : null),
    fromExcel: (val: string) =>
      options.find((it) => isNameEquals(it.label, val))?.value || null,
    options,
    dataType,
  };
}

export function getLookupCodeArrayMapping(
  options: { label: string; value: any }[],
  toExcelDefault = null,
  supportCreating = false,
  dataType: CellDataType = 'numbers',
): ITransferWithDataType {
  return <ITransferWithDataType>{
    toExcel: (val?: number[]) => {
      if (val?.length ?? 0 > 0) {
        return val!.map(id => {
          const label = options.find(opt => opt.value === id)?.label;
          return !label?.length ? toExcelDefault : label;
        })
          .filter(it => it != null)
          .join(',');
      } else {
        return '';
      }
    },
    fromExcel: (val: string) => {
      if (!val?.length) { return []; }
      const valArray = val.split(',').filter(it => it?.trim()?.length > 0);
      return valArray.map(it => {
        const optId = options.find(opt => isNameEquals(opt.label, it))?.value;
        if (optId == null && supportCreating) {
          const fakeItem = { label: it, value: getFakeId() };
          options.push(fakeItem);
          return fakeItem.value;
        } else {
          return optId;
        }
      }).filter(it => !!it);
    },
    options,
    showOptions: false,
    dataType,
  };
}

export function getNumberMapping(): ITransferWithDataType {
  return {
    toExcel: (val?: any) => {
      if (val == null) {
        return null;
      } else {
        return Number.isNaN(val) ? null : val;
      }
    },
    fromExcel: (val: any) => {
      if (val == null) return null;
      if (typeof val === 'number' || typeof val === 'bigint') {
        return val;
      } else {
        let num = Number.parseFloat(val);
        return Number.isNaN(num) ? null : num;
      }
    },
    dataType: 'number',
  };
}

export function getStringMapping(): ITransferWithDataType {
  return {
    toExcel: (val: any) => `${val || ''}`,
    fromExcel: (val: any) => `${val || ''}`.trim(),
    dataType: 'string',
  };
}

function getDefaultNumbersMapping(): ITransferWithDataType {
  return <ITransferWithDataType>{
    fromExcel: (val: string) => val?.length > 0 ? val.split(';') : [],
    toExcel: (val: any[]) => val.length > 0 ? val.join(';') : '',
    dataType: 'numbers',
  }
}

const yesConsts = ['yes', 'ok', 'true', 'allow', 'visible', 'y'];
const noConsts = ['no', 'false', 'forbidden', 'disable', 'hide', 'invisible', 'n'];
const nullConsts = ['mute'];

function isYesNoWord(val: string) {
  if (!val?.length) {
    return false;
  }
  return yesConsts.concat(noConsts).concat(nullConsts).includes(val.toLocaleLowerCase()?.trim());
}

export function getYesNoMapping(showBlank = false): ITransferWithDataType {
  return <ITransferWithDataType>{
    toExcel: (val?: boolean) =>
      val != null ? (val === true ? 'Yes' : 'No') : (showBlank ? '(blank)' : null),
    fromExcel: (val: boolean | string) => {
      if (val === true || val === false) return val;
      if (val == null) return null;
      if (typeof (val) != 'string') return null;
      if (yesConsts.includes(val.toLocaleLowerCase()?.trim())) {
        return true;
      }
      if (noConsts.includes(val.toLocaleLowerCase()?.trim())
      ) {
        return false;
      }
      return null;
    },
    options: yesNoOptions,
    dataType: 'boolean',
  };
}

export function getBooleanMapping(): ITransferWithDataType {
  return <ITransferWithDataType>{
    toExcel: (val?: boolean) =>
      val != null ? (val === true ? 'TRUE' : 'FALSE') : null,
    fromExcel: (val: boolean | string) => {
      if (val === true || val === false) return val;
      if (typeof (val) != 'string' || !val?.length || ['na', 'null'].includes(val.toLocaleLowerCase())) {
        return null;
      }
      if (
        ['yes', 'ok', 'true', 'allow', 'visible'].includes(
          val.toLocaleLowerCase()?.trim()
        )
      ) {
        return true;
      }
      if (
        [
          'no',
          'false',
          'forbidden',
          'disable',
          'mute',
          'hide',
          'invisible',
        ].includes(val.toLocaleLowerCase()?.trim())
      ) {
        return false;
      }
      return true;
    },
    options: yesNoOptions,
    dataType: 'boolean',
  };
}

export function getDateMapping(): ITransferWithDataType {
  return {
    toExcel: (val: string) => {
      return val;
    },
    fromExcel: (val: number | string | Date) => {
      try {
        if (!val) { return null; }
        //https://github.com/SheetJS/sheetjs/issues/1223
        if (typeof val === 'number') {
          const date = DateTime.fromMillis(Math.round((val - 25569) * 86400 * 1000));
          if (date.year > 3000 || date.year < 1000) {
            return null;
          }
          return date
            .toFormat('yyyy-MM-dd');
        }
        else if (Object.prototype.toString.call(val) === '[object Date]') {
          const date = DateTime.fromJSDate(val as Date, { zone: 'utc' });
          if (date.year > 3000 || date.year < 1000) {
            return null;
          }
          return date
            .toFormat('yyyy-MM-dd')
        }
        else if (typeof val === 'string') {
          return parseDateFromString(val);
        }
        else {
          return null;
        }
      } catch (ex) {
        return null;
      }
    },
    dataType: 'dateTime',
  };
}

function parseDateFromString(val: string) {
  let formats = ['yyyy-MM-dd', 'MM-dd-yyyy', 'yyyy-M-dd', 'yyyy.MM.dd', 'MM.dd.yyyy'];
  for (let format of formats) {
    let date = DateTime.fromFormat(val, format);
    if (date && date.toFormat('yyyy-MM-dd') !== 'Invalid DateTime') {
      return date.toFormat('yyyy-MM-dd');
    }
  }
  return null;
}

export function getDirectMapping(dataType: CellDataType = 'string'): ITransferWithDataType {
  return {
    toExcel: (val: any) => val,
    fromExcel: (val: any) => val,
    dataType,
  };
}

export function getDefaultMappingByType(dataType: CellDataType): ITransferWithDataType {
  switch (dataType) {
    case 'boolean': return getYesNoMapping();
    case 'dateTime': return getDateMapping();
    case 'number': return getNumberMapping();
    case 'string': return getStringMapping();
    case 'numbers': return getDefaultNumbersMapping();
    default: return getDirectMapping();
  }
}

function getDataTypeByValue(val: any): CellDataType {
  if (typeof val === 'bigint' || typeof val === 'number') {
    return 'number';
  } else if (typeof val === 'boolean') {
    return 'boolean';
  } else if (val != null && Object.prototype.toString.call(val) === '[object Date]') {
    return 'dateTime';
    // } else if (val != null && isYesNoWord(val)) {
    //   return 'boolean';
  } else if (Array.isArray(val)) {
    return 'numbers';
  }
  else {
    return 'string';
  }
}

export function getDefaultMappingByValueType(val: any): ITransferWithDataType {
  return getDefaultMapping(getDataTypeByValue(val));
}

export function getDefaultMapping(dataType?: CellDataType, val?: any): ITransferWithDataType {
  if (dataType != null) {
    return getDefaultMappingByType(dataType);
  } else {
    return getDefaultMappingByValueType(val);
  }
}

export function constructSheetColumnMetasFromObject(data: any): Pick<ISheetMetaDataWithRows, 'columns'> {
  let columns: ICellMetaData[] = [];
  if (data == null) {
    throw new Error('data can not be null');
  }
  Object.entries(data).forEach(([key, value]) => {
    let cell: ICellMetaData = {
      fieldName: key,
      header: key,
      ...getDefaultMappingByValueType(value)
    };
    columns.push(cell);
  })
  return { columns };
}


