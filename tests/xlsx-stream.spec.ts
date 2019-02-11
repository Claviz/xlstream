import { getXlsxStream } from '../src';

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
    // stream.resume();
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});
