import { CellDataType, ITransferWithDataType, ISheetMetaDataWithRows } from "./excel-model";
export declare const yesNoOptions: {
    label: string;
    value: boolean;
}[];
export declare function isNameEquals(original: string, other: string): boolean;
export declare function getLookupCodeMapping(options: {
    label: string;
    value: any;
}[], toExcelDefault?: null, dataType?: CellDataType): ITransferWithDataType;
export declare function getLookupCodeArrayMapping(options: {
    label: string;
    value: any;
}[], toExcelDefault?: null, supportCreating?: boolean, dataType?: CellDataType): ITransferWithDataType;
export declare function getNumberMapping(): ITransferWithDataType;
export declare function getStringMapping(): ITransferWithDataType;
export declare function getYesNoMapping(showBlank?: boolean): ITransferWithDataType;
export declare function getBooleanMapping(): ITransferWithDataType;
export declare function getDateMapping(): ITransferWithDataType;
export declare function getDirectMapping(dataType?: CellDataType): ITransferWithDataType;
export declare function getDefaultMappingByType(dataType: CellDataType): ITransferWithDataType;
export declare function getDefaultMappingByValueType(val: any): ITransferWithDataType;
export declare function getDefaultMapping(dataType?: CellDataType, val?: any): ITransferWithDataType;
export declare function constructSheetColumnMetasFromObject(data: any): Pick<ISheetMetaDataWithRows, 'columns'>;
//# sourceMappingURL=excel-helper.d.ts.map