export interface IXlsxStreamOptions {
    filePath: string;
    sheet: string | number | number[];
    withHeader?: boolean;
    ignoreEmpty?: boolean;
}

export interface IWorksheetOptions {
    filePath: string;
}
