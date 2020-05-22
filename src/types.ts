export interface IXlsxStreamOptions {
    filePath: string;
    sheet: string | number;
    withHeader?: boolean;
    ignoreEmpty?: boolean;
    fillMergedCells?: boolean;
}

export interface IXlsxStreamsOptions {
    filePath: string;
    sheets: {
        id: string | number;
        withHeader?: boolean;
        ignoreEmpty?: boolean;
        fillMergedCells?: boolean;
    }[];
}

export interface IWorksheetOptions {
    filePath: string;
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