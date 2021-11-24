import { getXlsxStream, getXlsxStreams, getWorksheets } from '../src';
import { Transform } from 'stream';
import { open } from 'fs';

it('reads XLSX file correctly', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/basic.xlsx',
        sheet: 0,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads empty XLSX file correctly', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/empty.xlsx',
        sheet: 0,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads XLSX file with header', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/with-header.xlsx',
        sheet: 0,
        withHeader: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads XLSX file with header values being dupicated', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/with-header-duplicated.xlsx',
        sheet: 0,
        withHeader: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('ignores empty rows', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/empty-rows.xlsx',
        sheet: 0,
        ignoreEmpty: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('does not ignore empty rows with data when ignoreEmpty is false', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/empty-rows.xlsx',
        sheet: 0,
        ignoreEmpty: false,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('adds empty rows not containing data when ignoreEmpty is false', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/empty-rows-missing.xlsx',
        sheet: 0,
        ignoreEmpty: false
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('gets worksheets', async (done) => {
    const sheets = await getWorksheets({
        filePath: './tests/assets/worksheets.xlsx',
    });
    expect(sheets).toEqual([
        { name: 'Sheet1', hidden: false, },
        { name: 'Sheet2', hidden: false, },
        { name: 'Sheet3', hidden: false, },
        { name: 'Sheet4', hidden: false, },
    ]);
    done();
});

it('gets worksheets with correct hidden info', async (done) => {
    const sheets = await getWorksheets({
        filePath: './tests/assets/hidden-sheet.xlsx',
    });
    expect(sheets).toEqual([
        { name: 'Sheet1', hidden: false, },
        { name: 'Sheet2', hidden: true, },
        { name: 'Sheet3', hidden: false, },
    ]);
    done();
});

it('get worksheets should fail gracefully for a corrupted xlsx', () => {
    return getWorksheets({
        filePath: './tests/assets/corrupted-file.xlsx',
    }).catch(e => expect(e).toMatch('Bad archive'));
});

it('gets worksheet by name, even if they are reordered', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/worksheets-reordered.xlsx',
        sheet: 'Sheet1',
        ignoreEmpty: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('gets worksheet by index, even if they are reordered', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/worksheets-reordered.xlsx',
        sheet: 1,
        ignoreEmpty: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it(`doesn't fail when empty row has custom height`, async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/empty-row-custom-height.xlsx',
        sheet: 0,
        ignoreEmpty: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it(`throws expected bad archive error`, async (done) => {
    const data: any = [];
    try {
        const stream = await getXlsxStream({
            filePath: './tests/assets/bad-archive.xlsx',
            sheet: 0,
        });
    } catch (err) {
        expect(err).toMatchSnapshot();
        done();
    }
});

it(`reads 2 sheets from XLSX file using generator`, async (done) => {
    const data: any = [];
    const generator = await getXlsxStreams({
        filePath: './tests/assets/worksheets-reordered.xlsx',
        sheets: [
            {
                id: 2,
                ignoreEmpty: true
            },
            {
                id: 'Sheet1',
                ignoreEmpty: true
            }
        ]
    });
    function processSheet(stream: Transform) {
        return new Promise<void>((resolve, reject) => {
            stream.on('data', (x: any) => data.push(x));
            stream.on('end', () => { resolve() });
        });
    }
    for await (const stream of generator) {
        await processSheet(stream);
    }
    expect(data).toMatchSnapshot();
    done();
});

it('fills merged cells with data', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/merged-cells.xlsx',
        sheet: 0,
        fillMergedCells: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('fills merged cells with data if header has merged cells', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/merged-cells-with-header.xlsx',
        sheet: 0,
        fillMergedCells: true,
        withHeader: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('correctly handles shared string if it has just one value in cell', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/shared-string-single-value.xlsx',
        sheet: 0,
        ignoreEmpty: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads XLSX file with header located in specific row', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/with-header-number.xlsx',
        sheet: 0,
        withHeader: 5,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads number values with leading zeroes correctly', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/zeroes.xlsx',
        sheet: 0,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('correctly reads file created with OpenXML (with `x` namespaces)', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/open-xml.xlsx',
        sheet: 0,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it(`applies custom number format`, async (done) => {
    const data: any = [];
    const generator = await getXlsxStreams({
        filePath: './tests/assets/custom-number-format.xlsx',
        sheets: [
            {
                id: 0,
            },
            {
                id: 1,
                numberFormat: 'excel',
            },
            {
                id: 2,
                numberFormat: {
                    47: 'mm ss'
                },
            }
        ]
    });
    function processSheet(stream: Transform) {
        return new Promise<void>((resolve, reject) => {
            stream.on('data', (x: any) => data.push(x));
            stream.on('end', () => { resolve() });
        });
    }
    for await (const stream of generator) {
        await processSheet(stream);
    }
    expect(data).toMatchSnapshot();
    done();
});

it('reads XLSX file with header values being undefined', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/multiple-undefined-columns-as-header.xlsx',
        sheet: 0,
        withHeader: true,
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('added encoding option support', async (done) => {
    const data: any = [];
    const stream = await getXlsxStream({
        filePath: './tests/assets/added-encoding-option-support.xlsx',
        sheet: 0,
        encoding: 'utf8'
    });
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});