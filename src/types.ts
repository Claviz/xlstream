export interface IXlsxStreamOptions {
    filePath: string;
    sheet: string | number;
    withHeader?: boolean;
    ignoreEmpty?: boolean;
    dontCloseZip?: boolean;
}

export interface IWorksheetOptions {
    filePath: string;
}

export interface IXlsxObj {
    options: IXlsxStreamOptions;
    zip: any;
    zipClosed: boolean;
    sheets: string[];
    formats: (string | number)[];
    strings: string[];
}
