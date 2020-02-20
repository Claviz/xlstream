export interface IXlsxStreamOptions {
    filePath: string;
    sheet: string | number;
    withHeader?: boolean;
    ignoreEmpty?: boolean;
}

export interface IXlsxStreamsOptions {
    filePath: string;
    sheets: {
        id: string | number;
        withHeader?: boolean;
        ignoreEmpty?: boolean;
    }[];
}

export interface IWorksheetOptions {
    filePath: string;
}
