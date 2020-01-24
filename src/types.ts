export interface IXlsxStreamOptions {
    filePath: string;
    sheet: string | number;
    withHeader?: boolean;
    ignoreEmpty?: boolean;
    skipRows?: Array<number>;
}

export interface IWorksheetOptions {
    filePath: string;
}
