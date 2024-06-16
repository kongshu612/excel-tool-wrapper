import { ISheetMetaData, IXlsxMetaData, IXlsxData } from './excel-model';
export declare function exportObjects(data: any[], fileName: string, showHeader: boolean, sheetName: string): Promise<void>;
export declare function exportObjects(data: {
    [key in string]: any[];
}, fileName: string, showHeader: boolean): Promise<void>;
export declare function exportObjects(fileMeta: IXlsxMetaData): Promise<void>;
export declare function importObjects(file: File | string): Promise<IXlsxData | null>;
export declare function importObjects(file: File | string, fileMeta: ISheetMetaData[]): Promise<IXlsxData | null>;
export declare function importObjects(file: File | string, data: {
    [key in string]: any;
}, compatiableMode: boolean): Promise<IXlsxData | null>;
//# sourceMappingURL=excel-import-export.d.ts.map