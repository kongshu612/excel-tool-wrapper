const cellDataTypes = ['boolean', 'string', 'number', 'dateTime', 'numbers'] as const;
export type CellDataType = typeof cellDataTypes[number];

export interface ICellDataValidation {
  options: { label: string; value: any, disabled?: boolean }[];
  affectedRowCount?: number;
}

export interface IXlsxTranfer {
  toExcel: (val: any) => any;
  fromExcel: (val: any) => any;
}

export type ITransferWithDataType = IXlsxTranfer & Required<Pick<ICellMetaData, 'dataType'>>;

export interface ICellMetaData extends Partial<IXlsxTranfer> {
  fieldName: string;
  dataValidations?: ICellDataValidation;
  required?: boolean;
  dataType?: CellDataType;
  header?: string;
  wch?: number;
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



export interface IXlsxMetaData {
  sheets: ISheetMetaDataWithRows[];
  fileName: string;
}


export type IXlsxData = { [key: string]: any[] }
