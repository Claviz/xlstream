export type numberFormatType = 'standard' | 'excel' | { [formatId in number | string]: string };

export interface IXlsxStreamOptions {
    filePath: string;
    sheet: string | number;
    withHeader?: boolean | number;
    ignoreEmpty?: boolean;
    fillMergedCells?: boolean;
    numberFormat?: numberFormatType;
    encoding?: BufferEncoding;
}

export interface IXlsxStreamsOptions {
    filePath: string;
    encoding?: BufferEncoding;
    sheets: {
        id: string | number;
        withHeader?: boolean | number;
        ignoreEmpty?: boolean;
        fillMergedCells?: boolean;
        numberFormat?: numberFormatType;
    }[];
}

export interface IWorksheetOptions {
    filePath: string;
    encoding?: BufferEncoding;
}

export interface IWorksheet {
    name: string;
    hidden: boolean;
}

export interface IMergedCellDictionary {
    [row: number]: {
        [column: string]: {
            parent: {
                column: string;
                row: number;
            };
            value: {
                raw: any;
                formatted: any;
            };
        }
    };
}