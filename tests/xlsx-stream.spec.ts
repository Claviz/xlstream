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

it('gets worksheet names list', async (done) => {
    const sheets = await getWorksheets({
        filePath: './tests/assets/worksheets.xlsx',
    });
    expect(sheets).toEqual([
        'Sheet1',
        'Sheet2',
        'Sheet3',
        'Sheet4',
    ]);
    done();
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
        return new Promise((resolve, reject) => {
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