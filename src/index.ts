import numfmt from 'numfmt';
import fclone from 'fclone';
import path from 'path';
import ssf from 'ssf';
import { Transform } from 'stream';
import { pipeline } from 'stream/promises';
import { pipeline as pipelineCallback } from 'stream';
import StreamZip, { ZipEntry } from 'node-stream-zip';
import SaxStream from 'sax-stream';
import rename from 'deep-rename-keys';
import {
    IMergedCellDictionary,
    IWorksheet,
    IWorksheetOptions,
    IXlsxStreamOptions,
    IXlsxStreamsOptions,
    numberFormatType
} from './types';


let currentSheetProcessedSize = 0;
let currentSheetSize = 0;

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
}


function applyHeaderToObj(obj: any, header: any) {
    if (!header || !header.length) {
        return obj;
    }
    const newObj: { [key: string]: any } = {};
    for (const columnName of Object.keys(obj)) {
        const index = lettersToNumber(columnName) - 1;
        newObj[header[index] || `[${columnName}]`] = obj[columnName];
    }
    return newObj;
}

function getFilledHeader(arr: any[], header: any[]) {
    if (!header || !header.length) {
        return header;
    }
    const filledHeader = [];
    for (let i = 0; i < Math.max(arr.length, header.length); i++) {
        filledHeader.push(
            header[i] ? header[i] : `[${numbersToLetter(i + 1)}]`
        );
    }
    return filledHeader;
}

function fillMergedCells(dict: IMergedCellDictionary, currentRowName: any, arr: any, obj: any, formattedArr: any, formattedObj: any) {
    for (const columnName of Object.keys(dict[currentRowName])) {
        const parentCell = dict[currentRowName][columnName].parent;
        const index = lettersToNumber(columnName) - 1;
        arr[index] = obj[columnName] = dict[parentCell.row][parentCell.column].value.raw;
        formattedArr[index] = formattedObj[columnName] = dict[parentCell.row][parentCell.column].value.formatted;
    }
}

function formatNumericValue(attr: string, value: any) {
    if (attr === 'inlineStr' || attr === 's') {
        return value;
    }
    return isNaN(value) ? value : Number(value);
}

function getTransform(formats: (string | number)[], strings: string[], dict?: IMergedCellDictionary, withHeader?: boolean | number, ignoreEmpty?: boolean, numberFormat?: numberFormatType) {
    let lastReceivedRow = 0;
    let header: any[] = [];

    return new Transform({
        objectMode: true,
        transform(chunk, encoding, done) {
            let arr: any[] = [];
            let formattedArr = [];
            let obj: any = {};
            let formattedObj: any = {};
            const record = rename(fclone(chunk.record), (key: string) => {
                const keySplit = key.split(':');
                const tag = keySplit.length === 2 ? keySplit[1] : key;
                return tag;
            });
            const children = record.children ? record.children.c.length ? record.children.c : [record.children.c] : [];
            const nextRow = record.attribs ? parseInt(record.attribs.r) : lastReceivedRow + 1;
            if (!ignoreEmpty) {
                const emptyRowCount = nextRow - lastReceivedRow - 1;
                for (let i = 0; i < emptyRowCount; i++) {
                    this.push({
                        raw: {
                            obj: {},
                            arr: []
                        },
                        formatted: {
                            obj: {},
                            arr: []
                        },
                        header: getFilledHeader(arr, header),
                        processedSheetSize: currentSheetProcessedSize,
                        totalSheetSize: currentSheetSize,
                    })
                }
            }
            lastReceivedRow = nextRow;
            for (let i = 0; i < children.length; i++) {
                const ch = children[i];
                if (ch.children) {
                    let value: any;
                    const type = ch.attribs?.t;
                    const columnName = ch.attribs?.r;
                    const formatId = ch.attribs?.s ? Number(ch.attribs.s) : 0;
                    if (type === 'inlineStr') {
                        value = ch.children.is.children.t.value;
                    } else {
                        value = ch.children.v.value;
                        if (type === 's') {
                            value = strings[value];
                        }
                    }
                    value = formatNumericValue(type, value);
                    let column = columnName ? columnName.replace(/[0-9]/g, '') : numbersToLetter(i + 1);
                    const index = lettersToNumber(column) - 1;
                    if (dict?.[lastReceivedRow]?.[column]) {
                        dict[lastReceivedRow][column].value.raw = value;
                    }
                    arr[index] = value;
                    obj[column] = value;
                    if (formatId) {
                        let numFormat = formats[formatId];
                        if (numberFormat && numberFormat === 'excel' && typeof numFormat === 'number' && excelNumberFormat[numFormat]) {
                            numFormat = excelNumberFormat[numFormat];
                        } else if (numberFormat && typeof numberFormat === 'object') {
                            numFormat = numberFormat[numFormat];
                        }
                        if (typeof numFormat === 'string') {
                            numFormat = numFormat.replace(/[Hh]{2}/g, 'hh');
                            value = numfmt.format(numFormat, value);
                        } else {
                            value = ssf.format(numFormat, value);
                        }
                        value = formatNumericValue(type, value);
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
            if (((typeof withHeader === 'number' && withHeader === lastReceivedRow - 1) || (typeof withHeader !== 'number' && withHeader)) && !header.length) {
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
                    header: getFilledHeader(arr, header),
                    processedSheetSize: currentSheetProcessedSize,
                    totalSheetSize: currentSheetSize,
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
                        header: getFilledHeader(arr, header),
                        processedSheetSize: currentSheetProcessedSize,
                        totalSheetSize: currentSheetSize,
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
        encoding: options.encoding,
        sheets: [{
            id: options.sheet,
            withHeader: options.withHeader,
            ignoreEmpty: options.ignoreEmpty,
            fillMergedCells: options.fillMergedCells,
            numberFormat: options.numberFormat,
        }]
    });
    const stream = await generator.next();
    return stream.value;
}

export async function* getXlsxStreams(options: IXlsxStreamsOptions): AsyncGenerator<Transform> {
    const sheets: { relsId: string; name: string; }[] = [];
    const rels: { [id: string]: string } = {};
    const numberFormats: any = {};
    const formats: (string | number)[] = [];
    const strings: string[] = [];
    const zip = new StreamZip.async({
        file: options.filePath,
        storeEntries: true
    });
    let zipEntries: { [name: string]: ZipEntry } = {};
    let currentSheetIndex = 0;

    async function setupGenericData() {
        async function processSharedStrings(numberFormats: any, formats: (string | number)[]) {
            for (let i = 0; i < formats.length; i++) {
                const format = numberFormats[formats[i]];
                if (format) {
                    formats[i] = format;
                }
            }

            if (!zipEntries['xl/sharedStrings.xml']) {
                return;
            }

            const zipStream = await zip.stream('xl/sharedStrings.xml');

            if (options.encoding) {
                zipStream.setEncoding(options.encoding);
            }

            const saxStream = SaxStream({
                strict: true,
                tag: ['x:si', 'si']
            });

            await pipeline(
                zipStream,
                saxStream,
                async function* (elements) {
                    for await (const element of elements) {
                        const record = element.record;
                        if (record.children.t) {
                            strings.push(record.children.t.value);
                        } else if (!record.children.r.length) {
                            strings.push(record.children.r.children.t.value);
                        } else {
                            let str = '';
                            for (let i = 0; i < record.children.r.length; i++) {
                                str += record.children.r[i].children.t.value;
                            }
                            strings.push(str);
                        }
                    }
                }
            );
        }

        async function processStyles() {
            if (!zipEntries['xl/styles.xml']) {
                return;
            }

            const zipStream = await zip.stream(`xl/styles.xml`);

            if (options.encoding) {
                zipStream.setEncoding(options.encoding);
            }

            const saxStream = SaxStream({
                strict: true,
                tag: ['x:cellXfs', 'x:numFmts', 'cellXfs', 'numFmts']
            });

            await pipeline(
                zipStream,
                saxStream,
                async function* (elements) {
                    for await (const element of elements) {
                        if ((element.tag === 'numFmts' || element.tag === 'x:numFmts') && element.record.children) {
                            let numFmtField = element.record.children['x:numFmt'] ? 'x:numFmt' : 'numFmt';
                            const children = element.record.children[numFmtField].length ? element.record.children[numFmtField] : [element.record.children[numFmtField]];
                            for (let i = 0; i < children.length; i++) {
                                numberFormats[Number(children[i].attribs.numFmtId)] = children[i].attribs.formatCode;
                            }
                        } else if ((element.tag === 'cellXfs' || element.tag === 'x:cellXfs') && element.record.children) {
                            const xfField = element.record.children['x:xf'] ? 'x:xf' : 'xf';
                            for (let i = 0; i < element.record.children[xfField].length; i++) {
                                const ch = element.record.children[xfField][i];
                                if (ch.attribs?.numFmtId) {
                                    formats[i] = ch.attribs?.numFmtId ? Number(ch.attribs?.numFmtId) : '';
                                }
                            }
                        }
                    }
                },
            );
        }

        async function processWorkbook() {
            if (!zipEntries['xl/workbook.xml']) {
                return;
            }

            const zipStream = await zip.stream('xl/workbook.xml');
            const saxStream = SaxStream({
                strict: true,
                tag: ['x:sheet', 'sheet']
            })

            if (options.encoding) {
                zipStream.setEncoding(options.encoding);
            }

            await pipeline(
                zipStream,
                saxStream,
                async function* (elements) {
                    for await (const element of elements) {
                        const attribs = element.record.attribs;
                        sheets.push({ name: attribs.name, relsId: attribs['r:id'] });
                    }
                }
            )
        }

        async function getRels() {
            if (!zipEntries['xl/_rels/workbook.xml.rels']) {
                return;
            }

            const zipStream = await zip.stream('xl/_rels/workbook.xml.rels');
            const saxStream = SaxStream({
                strict: true,
                tag: ['x:Relationship', 'Relationship']
            });

            if (options.encoding) {
                zipStream.setEncoding(options.encoding);
            }

            await pipeline(
                zipStream,
                saxStream,
                async function* (elements) {
                    for await (const element of elements) {
                        rels[element.record.attribs.Id] = path.basename(element.record.attribs.Target);
                    }
                }
            );
        }

        zipEntries = await zip.entries();

        await getRels();
        await processWorkbook();
        await processStyles();
        await processSharedStrings(numberFormats, formats);
    }

    async function getMergedCellDictionary(sheetFileName: string) {
        const zipStream = await zip.stream(`xl/worksheets/${sheetFileName}`);
        const saxStream = SaxStream({
            strict: true,
            tag: ['x:mergeCell', 'mergeCell']
        });

        if (options.encoding) {
            zipStream.setEncoding(options.encoding);
        }

        const dict: IMergedCellDictionary = {};

        await pipeline(
            zipStream,
            saxStream,
            async function* (elements) {
                for await (const element of elements) {
                    const record = element.record;
                    const mergedCellRange: string = record.attribs.ref;
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
                }
            }
        );

        return dict;
    }

    async function getSheetTransform(sheetFileName: string, withHeader?: boolean | number, ignoreEmpty?: boolean, fillMergedCells?: boolean, numberFormat?: numberFormatType) {
        let dict: IMergedCellDictionary | undefined;
        if (fillMergedCells) {
            dict = await getMergedCellDictionary(sheetFileName);
        }

        const sheetFullFileName = `xl/worksheets/${sheetFileName}`;
        const zipStream = await zip.stream(sheetFullFileName);

        if (options.encoding) {
            zipStream.setEncoding(options.encoding);
        }

        currentSheetProcessedSize = 0;
        currentSheetSize = zipEntries[sheetFullFileName].size;

        const sheetTransformStream = getTransform(formats, strings, dict, withHeader, ignoreEmpty, numberFormat);

        pipelineCallback(
            zipStream,
            async function* (chunks) {
                for await (const chunk of chunks) {
                    currentSheetProcessedSize += chunk.length;

                    yield chunk;
                }
            },
            SaxStream({
                strict: true,
                tag: ['x:row', 'row']
            }),
            sheetTransformStream,
            () => {
                if (currentSheetIndex + 1 === options.sheets.length) {
                    zip.close();
                }
            }
        )

        return sheetTransformStream;
    }

    await setupGenericData();

    for (currentSheetIndex = 0; currentSheetIndex < options.sheets.length; currentSheetIndex++) {
        const sheet = options.sheets[currentSheetIndex];
        const id = sheet.id;
        let sheetIndex = 0;
        if (typeof id === 'number') {
            sheetIndex = id;
        } else if (typeof id === 'string') {
            sheetIndex = sheets.findIndex(x => x.name === id);
        }
        const sheetFileName = rels[sheets[sheetIndex].relsId];
        const transform = await getSheetTransform(sheetFileName, sheet.withHeader, sheet.ignoreEmpty, sheet.fillMergedCells, sheet.numberFormat);

        yield transform;
    }
}

export async function getWorksheets(options: IWorksheetOptions) {
    const sheets: IWorksheet[] = [];

    const zip = new StreamZip.async({
        file: options.filePath,
        storeEntries: true,
    });

    const zipStream = await zip.stream('xl/workbook.xml')


    if (options.encoding) {
        zipStream.setEncoding(options.encoding);
    }

    const saxStream = SaxStream({
        strict: true,
        tag: ['x:sheet', 'sheet'],
    });

    await pipeline(
        zipStream,
        saxStream,
        async function* (elements) {
            for await (const element of elements) {
                sheets.push({
                    name: element.record.attribs.name,
                    hidden: Boolean(element.record.attribs.state && element.record.attribs.state === 'hidden'),
                });
            }
        }
    );

    return sheets;
}

export const excelNumberFormat: { [format: number]: string } = {
    14: 'm/d/yyyy',
    22: 'm/d/yyyy h:mm',
    37: '#,##0_);(#,##0)',
    38: '#,##0_);[Red](#,##0)',
    39: '#,##0.00_);(#,##0.00)',
    40: '#,##0.00_);[Red](#,##0.00)',
    47: 'mm:ss.0',
}