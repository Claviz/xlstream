import ssf from 'ssf';
import { Transform } from 'stream';
import { ReadStream } from 'tty';

import { IMergedCellDictionary, IWorksheetOptions, IXlsxStreamOptions, IXlsxStreamsOptions, IWorksheet } from './types';

const StreamZip = require('node-stream-zip');
const saxStream = require('sax-stream');

function lettersToNumber(letters: string) {
    return letters.split('').reduce((r, a) => r * 26 + parseInt(a, 36) - 9, 0);
}

function numbersToLetter(number: number) {
    let colName = '';
    let dividend = Math.floor(Math.abs(number));
    let rest: number;

    while (dividend > 0) {
        rest = (dividend - 1) % 26;
        colName = String.fromCharCode(65 + rest) + colName;
        dividend = parseInt(`${(dividend - rest) / 26}`);
    }
    return colName;
};


function applyHeaderToObj(obj: any, header: any) {
    if (!header || !header.length) {
        return obj;
    }
    const newObj: { [key: string]: any } = {};
    for (const columnName of Object.keys(obj)) {
        const index = lettersToNumber(columnName) - 1;
        newObj[header[index]] = obj[columnName];
    }
    return newObj;
}

function fillMergedCells(dict: IMergedCellDictionary, currentRowName: any, arr: any, obj: any, formattedArr: any, formattedObj: any) {
    for (const columnName of Object.keys(dict[currentRowName])) {
        const parentCell = dict[currentRowName][columnName].parent;
        const index = lettersToNumber(columnName) - 1;
        arr[index] = obj[columnName] = dict[parentCell.row][parentCell.column].value.raw;
        formattedArr[index] = formattedObj[columnName] = dict[parentCell.row][parentCell.column].value.formatted;
    }
}

function getTransform(formats: (string | number)[], strings: string[], dict?: IMergedCellDictionary, withHeader?: boolean, ignoreEmpty?: boolean) {
    let lastReceivedRow: number;
    let header: any[] = [];
    return new Transform({
        objectMode: true,
        transform(chunk, encoding, done) {
            let arr: any[] = [];
            let formattedArr = [];
            let obj: any = {};
            let formattedObj: any = {};
            const children = chunk.children ? chunk.children.c.length ? chunk.children.c : [chunk.children.c] : [];
            lastReceivedRow = chunk.attribs.r;
            for (let i = 0; i < children.length; i++) {
                const ch = children[i];
                if (ch.children) {
                    let value: any;
                    if (ch.attribs.t === 'inlineStr') {
                        value = ch.children.is.children.t.value;
                    } else {
                        value = ch.children.v.value;
                        if (ch.attribs.t === 's') {
                            value = strings[value];
                        }
                    }
                    value = isNaN(value) ? value : Number(value);
                    let column = ch.attribs.r.replace(/[0-9]/g, '');
                    const index = lettersToNumber(column) - 1;
                    if (dict?.[lastReceivedRow]?.[column]) {
                        dict[lastReceivedRow][column].value.raw = value;
                    }
                    arr[index] = value;
                    obj[column] = value;
                    const formatId = ch.attribs.s ? Number(ch.attribs.s) : 0;
                    if (formatId) {
                        value = ssf.format(formats[formatId], value);
                        value = isNaN(value) ? value : Number(value);
                    }
                    if (dict?.[lastReceivedRow]?.[column]) {
                        dict[lastReceivedRow][column].value.formatted = value;
                    }
                    formattedArr[index] = value;
                    formattedObj[column] = value;
                }
            }
            if (dict?.[lastReceivedRow]) {
                fillMergedCells(dict, lastReceivedRow, arr, obj, formattedArr, formattedObj);
            }
            if (withHeader && !header.length) {
                for (let i = 0; i < arr.length; i++) {
                    const hasDuplicate = arr.filter(x => x === arr[i]).length > 1;
                    header[i] = hasDuplicate ? `[${numbersToLetter(i + 1)}] ${arr[i]}` : arr[i];
                }
                done();
            } else {
                done(undefined, ignoreEmpty && !arr.length ? null : {
                    raw: {
                        obj: applyHeaderToObj(obj, header),
                        arr
                    },
                    formatted: {
                        obj: applyHeaderToObj(formattedObj, header),
                        arr: formattedArr,
                    },
                    header,
                });
            }
        },
        flush(callback) {
            if (dict) {
                const unprocessedRows = Object.keys(dict).map(x => Number(x)).filter(x => x > lastReceivedRow);
                for (const unprocessedRow of unprocessedRows) {
                    let arr: any[] = [];
                    let formattedArr: any[] = [];
                    let obj: any = {};
                    let formattedObj: any = {};
                    fillMergedCells(dict, unprocessedRow, arr, obj, formattedArr, formattedObj);
                    this.push((ignoreEmpty && !arr.length) ? null : {
                        raw: {
                            obj: applyHeaderToObj(obj, header),
                            arr
                        },
                        formatted: {
                            obj: applyHeaderToObj(formattedObj, header),
                            arr: formattedArr,
                        },
                        header,
                    });
                }
            }
            callback();
        }
    })
}

export async function getXlsxStream(options: IXlsxStreamOptions): Promise<Transform> {
    const generator = getXlsxStreams({
        filePath: options.filePath,
        sheets: [{
            id: options.sheet,
            withHeader: options.withHeader,
            ignoreEmpty: options.ignoreEmpty,
            fillMergedCells: options.fillMergedCells,
        }]
    });
    const stream = await generator.next();
    return stream.value;
}

export async function* getXlsxStreams(options: IXlsxStreamsOptions): AsyncGenerator<Transform> {
    const sheets: string[] = [];
    const numberFormats: any = {};
    const formats: (string | number)[] = [];
    const strings: string[] = [];
    const zip = new StreamZip({
        file: options.filePath,
        storeEntries: true
    });
    let currentSheetIndex = 0;
    function setupGenericData() {
        return new Promise((resolve, reject) => {
            function processSharedStrings(numberFormats: any, formats: (string | number)[]) {
                for (let i = 0; i < formats.length; i++) {
                    const format = numberFormats[formats[i]];
                    if (format) {
                        formats[i] = format;
                    }
                }
                zip.stream('xl/sharedStrings.xml', (err: any, stream: ReadStream) => {
                    if (stream) {
                        stream.pipe(saxStream({
                            strict: true,
                            tag: 'si'
                        })).on('data', (x: any) => {
                            if (x.children.t) {
                                strings.push(x.children.t.value);
                            } else {
                                let str = '';
                                for (let i = 0; i < x.children.r.length; i++) {
                                    const ch = x.children.r[i].children;
                                    str += ch.t.value;
                                }
                                strings.push(str);
                            }
                        });
                        stream.on('end', () => {
                            resolve();
                        });
                    } else {
                        resolve();
                    }
                });
            }

            function processStyles() {
                zip.stream(`xl/styles.xml`, (err: any, stream: ReadStream) => {
                    stream.pipe(saxStream({
                        strict: true,
                        tag: ['cellXfs', 'numFmts']
                    })).on('data', (x: any) => {
                        if (x.tag === 'numFmts' && x.record.children) {
                            const children = x.record.children.numFmt.length ? x.record.children.numFmt : [x.record.children.numFmt];
                            for (let i = 0; i < children.length; i++) {
                                numberFormats[Number(children[i].attribs.numFmtId)] = children[i].attribs.formatCode;
                            }
                        } else if (x.tag === 'cellXfs' && x.record.children) {
                            for (let i = 0; i < x.record.children.xf.length; i++) {
                                const ch = x.record.children.xf[i];
                                formats[i] = Number(ch.attribs.numFmtId);
                            }
                        }
                    });
                    stream.on('end', () => {
                        processSharedStrings(numberFormats, formats);
                    });
                });
            }

            function processWorkbook() {
                zip.stream('xl/workbook.xml', (err: any, stream: ReadStream) => {
                    stream.pipe(saxStream({
                        strict: true,
                        tag: 'sheet'
                    })).on('data', (x: any) => {
                        const attribs = x.attribs;
                        sheets.push(attribs.name);
                    });
                    stream.on('end', () => {
                        processStyles();
                    });
                });
            }

            zip.on('ready', () => {
                processWorkbook();
            });
            zip.on('error', (err: any) => {
                reject(new Error(err));
            });
        });
    }
    function getMergedCellDictionary(sheetId: string) {
        return new Promise<IMergedCellDictionary>((resolve, reject) => {
            zip.stream(`xl/worksheets/sheet${sheetId}.xml`, (err: any, stream: ReadStream) => {
                const dict: IMergedCellDictionary = {};
                const readStream = stream
                    .pipe(saxStream({
                        strict: true,
                        tag: 'mergeCell'
                    }));
                readStream.on('end', () => {
                    resolve(dict);
                });
                readStream.on('data', (a: any) => {
                    const mergedCellRange: string = a.attribs.ref;
                    const mergedCellRangeSplit = mergedCellRange.split(':');
                    const mergedCellRangeStart = mergedCellRangeSplit[0];
                    const mergedCellRangeEnd = mergedCellRangeSplit[1];
                    let columnLetterStart = mergedCellRangeStart.replace(/[0-9]/g, '');
                    let columnNumberStart = lettersToNumber(columnLetterStart);
                    let rowNumberStart = Number(mergedCellRangeStart.replace(columnLetterStart, ''));
                    let columnLetterEnd = mergedCellRangeEnd.replace(/[0-9]/g, '');
                    let columnNumberEnd = lettersToNumber(columnLetterEnd);
                    let rowNumberEnd = Number(mergedCellRangeEnd.replace(columnLetterEnd, ''));
                    for (let rowNumber = rowNumberStart; rowNumber <= rowNumberEnd; rowNumber++) {
                        for (let columnNumber = columnNumberStart; columnNumber <= columnNumberEnd; columnNumber++) {
                            const columnLetter = numbersToLetter(columnNumber);
                            if (!dict[rowNumber]) {
                                dict[rowNumber] = {};
                            }
                            dict[rowNumber][columnLetter] = {
                                parent: {
                                    column: columnLetterStart,
                                    row: rowNumberStart,
                                },
                                value: { formatted: null, raw: null },
                            }
                        }
                    }
                });
                readStream.resume();
            });
        });
    }
    async function getSheetTransform(sheetId: string, withHeader?: boolean, ignoreEmpty?: boolean, fillMergedCells?: boolean) {
        let dict: IMergedCellDictionary | undefined;
        if (fillMergedCells) {
            dict = await getMergedCellDictionary(sheetId);
        }
        return new Promise<Transform>((resolve, reject) => {
            zip.stream(`xl/worksheets/sheet${sheetId}.xml`, (err: any, stream: ReadStream) => {
                const readStream = stream
                    .pipe(saxStream({
                        strict: true,
                        tag: 'row'
                    }))
                    .pipe(getTransform(formats, strings, dict, withHeader, ignoreEmpty));
                readStream.on('end', () => {
                    if (currentSheetIndex + 1 === options.sheets.length) {
                        zip.close();
                    }
                });
                resolve(readStream);
            });
        });
    }
    await setupGenericData();
    for (currentSheetIndex = 0; currentSheetIndex < options.sheets.length; currentSheetIndex++) {
        const id = options.sheets[currentSheetIndex].id;
        let sheetId: string = '';
        if (typeof id === 'number') {
            sheetId = `${id + 1}`;
        } else if (typeof id === 'string') {
            sheetId = `${sheets.indexOf(id) + 1}`;
        }
        const transform = await getSheetTransform(sheetId, options.sheets[currentSheetIndex].withHeader, options.sheets[currentSheetIndex].ignoreEmpty, options.sheets[currentSheetIndex].fillMergedCells);

        yield transform;
    }
}

export function getWorksheets(options: IWorksheetOptions) {
    return new Promise<IWorksheet[]>((resolve, reject) => {
        function processWorkbook() {
            zip.stream('xl/workbook.xml', (err: any, stream: ReadStream) => {
                stream.pipe(saxStream({
                    strict: true,
                    tag: 'sheet'
                })).on('data', (x: any) => {
                    sheets.push({
                        name: x.attribs.name,
                        hidden: x.attribs.state && x.attribs.state === 'hidden' ? true : false,
                    });
                });
                stream.on('end', () => {
                    zip.close();
                    resolve(sheets);
                });
            });
        }

        let sheets: IWorksheet[] = [];
        const zip = new StreamZip({
            file: options.filePath,
            storeEntries: true,
        });
        zip.on('ready', () => {
            processWorkbook();
        });
    });
}
