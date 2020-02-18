import ssf from 'ssf';
import { Transform } from 'stream';
import { ReadStream } from 'tty';

import { IXlsxStreamOptions, IWorksheetOptions } from './types';

const StreamZip = require('node-stream-zip');
const saxStream = require('sax-stream');

function lettersToNumber(letters: string) {
    return letters.split('').reduce((r, a) => r * 26 + parseInt(a, 36) - 9, 0);
}

export function getXlsxStream(options: IXlsxStreamOptions) {
    return new Promise<Transform | Transform[]>((resolve, reject) => {

        function getTransform(formats: (string | number)[], strings: string[], header: string[]) {
            return new Transform({
                objectMode: true,
                transform: (chunk, encoding, done) => {
                    let arr = [];
                    let formattedArr = [];
                    let obj: any = {};
                    let formattedObj: any = {};
                    let parsingHeader = false;
                    const children = chunk.children ? chunk.children.c.length ? chunk.children.c : [chunk.children.c] : [];
                    for (let i = 0; i < children.length; i++) {
                        const ch = children[i];
                        if (ch.children) {
                            let value = ch.children.v.value;
                            if (ch.attribs.t === 's') {
                                value = strings[value];
                            }
                            value = isNaN(value) ? value : Number(value);
                            let column = ch.attribs.r.replace(/[0-9]/g, '');
                            const index = lettersToNumber(column) - 1;
                            if (options.withHeader) {
                                if (!parsingHeader && header.length) {
                                    column = header[index];
                                } else {
                                    header[index] = value;
                                    parsingHeader = true;
                                }
                            }
                            arr[index] = value;
                            obj[column] = value;
                            const formatId = ch.attribs.s ? Number(ch.attribs.s) : 0;
                            if (formatId) {
                                value = ssf.format(formats[formatId], value);
                                value = isNaN(value) ? value : Number(value);
                            }
                            formattedArr[index] = value;
                            formattedObj[column] = value;
                        }
                    }
                    done(undefined, parsingHeader || (options.ignoreEmpty && !arr.length) ? null : {
                        raw: {
                            obj,
                            arr
                        },
                        formatted: {
                            obj: formattedObj,
                            arr: formattedArr,
                        },
                        header,
                    });
                }
            })
        }

        async function processSheets(sheetIds: (string | number | number[]), formats: (string | number)[], strings: string[]) {
            if (Array.isArray(sheetIds)) {
                const readStreams: Transform[] = [];
                let streamsActive = 0;
                sheetIds.forEach(function (idx, n) {
                    zip.stream(`xl/worksheets/sheet${idx + 1}.xml`, (err: any, stream: ReadStream) => {
                        readStreams[n] = stream
                            .pipe(saxStream({
                            strict: true,
                            tag: 'row'
                        }))
                        .pipe(getTransform(formats, strings, []));
                        streamsActive++;
                        stream.on('end', () => {
                            streamsActive--;
                            if (!streamsActive) {
                                zip.close();
                            }
                        });
                    });
                });
                let streamsReady = false;
                while (!streamsReady) {
                    streamsReady = true;
                    sheetIds.forEach(function (idx, n) {
                        if (typeof readStreams[n] === 'undefined') {
                            streamsReady = false;
                        }
                    });
                    let p = new Promise(function (resolve) {
                        setTimeout( function() { resolve(true); }, 10 );
                    });
                    await p;
                }
                resolve(readStreams);
            } else {
                zip.stream(`xl/worksheets/sheet${sheetIds}.xml`, (err: any, stream: ReadStream) => {
                    const readStream = stream
                        .pipe(saxStream({
                            strict: true,
                            tag: 'row'
                        }))
                        .pipe(getTransform(formats, strings, []));
                    stream.on('end', () => {
                        zip.close();
                    });
                    resolve(readStream);
                });
            }
        }

        function processSharedStrings(sheetIds: (string | number | number[]), numberFormats: any, formats: (string | number)[]) {
            const strings: string[] = [];
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
                        processSheets(sheetIds, formats, strings);
                    });
                } else {
                    processSheets(sheetIds, formats, strings);
                }
            });
        }

        function processStyles(sheetIds: (string | number | number[])) {
            zip.stream(`xl/styles.xml`, (err: any, stream: ReadStream) => {
                const numberFormats: any = {};
                const formats: (string | number)[] = [];
                stream.pipe(saxStream({
                    strict: true,
                    tag: ['cellXfs', 'numFmts']
                })).on('data', (x: any) => {
                    if (x.tag === 'numFmts') {
                        const children = x.record.children.numFmt.length ? x.record.children.numFmt : [x.record.children.numFmt];
                        for (let i = 0; i < children.length; i++) {
                            numberFormats[Number(children[i].attribs.numFmtId)] = children[i].attribs.formatCode;
                        }
                    } else if (x.tag === 'cellXfs') {
                        for (let i = 0; i < x.record.children.xf.length; i++) {
                            const ch = x.record.children.xf[i];
                            formats[i] = Number(ch.attribs.numFmtId);
                        }
                    }
                });
                stream.on('end', () => {
                    processSharedStrings(sheetIds, numberFormats, formats);
                });
            });
        }

        function processWorkbook() {
            zip.stream('xl/workbook.xml', (err: any, stream: ReadStream) => {
                const sheets: string[] = [];
                stream.pipe(saxStream({
                    strict: true,
                    tag: 'sheet'
                })).on('data', (x: any) => {
                    const attribs = x.attribs;
                    sheets.push(attribs.name);
                });
                stream.on('end', () => {
                    if (typeof options.sheet === 'number') {
                        processStyles(`${options.sheet + 1}`);
                    } else if (typeof options.sheet === 'string') {
                        processStyles(`${sheets.indexOf(options.sheet) + 1}`);
                    } else if (Array.isArray(options.sheet)) {
                        processStyles(options.sheet);
                    }
                });
            });
        }

        const zip = new StreamZip({
            file: options.filePath,
            storeEntries: true
        });
        zip.on('ready', () => {
            processWorkbook();
        });
        zip.on('error', (err: any) => {
            reject(new Error(err));
        });
    });
}

export function getWorksheets(options: IWorksheetOptions) {
    return new Promise<string[]>((resolve, reject) => {
        function processWorkbook() {
            zip.stream('xl/workbook.xml', (err: any, stream: ReadStream) => {
                stream.pipe(saxStream({
                    strict: true,
                    tag: 'sheet'
                })).on('data', (x: any) => {
                    sheets.push(x.attribs.name);
                });
                stream.on('end', () => {
                    zip.close();
                    resolve(sheets);
                });
            });
        }

        let sheets: string[] = [];
        const zip = new StreamZip({
            file: options.filePath,
            storeEntries: true,
        });
        zip.on('ready', () => {
            processWorkbook();
        });
    });
}