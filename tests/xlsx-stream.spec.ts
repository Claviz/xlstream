import { getXlsxStream, getWorksheets } from '../src';

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