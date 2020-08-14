[![Build Status](https://travis-ci.org/Claviz/xlstream.svg?branch=master)](https://travis-ci.org/Claviz/xlstream)
[![codecov](https://codecov.io/gh/Claviz/xlstream/branch/master/graph/badge.svg)](https://codecov.io/gh/Claviz/xlstream)
![npm](https://img.shields.io/npm/v/xlstream.svg)

# xlstream

Memory-efficiently turns XLSX file into a [transform stream](https://nodejs.org/api/stream.html#stream_duplex_and_transform_streams) with all its benefits.

* Stream is **pausable**.
* Emits all **default events** (`data`, `end`, etc.)
* Returns **header**, **raw** and **formatted** row data in just one `data` event.
* Maintains desirable behavior of **merged cells**.

## Installation
```
npm install xlstream
```

## Example
`source.xlsx` contents:

| A     | B   |
|-------|-----|
| hello | 123 |

Where `123` is a `123.123` number formatted to be rounded to integer.

Script:
```javascript
const { getXlsxStream } = require('xlstream');

(async () => {
    const stream = await getXlsxStream({
        filePath: './source.xlsx',
        sheet: 0,
    });
    stream.on('data', x => console.log(x));
})();
```
Result:
```JSON
{ 
    "raw": { 
        "obj": { "A": "hello", "B": 123.123 }, 
        "arr": [ "hello", 123.123 ] 
    },
    "formatted": { 
        "obj": { "A": "hello", "B": 123 }, 
        "arr": [ "hello", 123 ] 
    },
    "header": []
}
```

## getXlsxStream
Returns transform stream of the sheet.

### Options

| option          | type                  | description                                                                                                                                                                                                                                                               |
|-----------------|-----------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| filePath        | `string`              | Path to the XLSX file                                                                                                                                                                                                                                                     |
| sheet           | `string` or `number`  | If `string` is passed, finds sheet by it's name. If `number`, finds sheet by it's index.                                                                                                                                                                                  |
| withHeader      | `boolean` or `number` | If `true`, column names will be taken from the first sheet row. If duplicated header name is found, column name will be prepended with column letter to maintain uniqueness. 0-based row location can be passed to this option if header is not located on the first row. |
| ignoreEmpty     | `boolean`             | If `true`, empty rows won't be emitted.                                                                                                                                                                                                                                   |
| fillMergedCells | `boolean`             | If `true`, merged cells will have the same value (by default, only the first cell of merged cells is filled with value). **Warning!** Enabling this feature may increase streaming time because file must be processed to detect merged cells before actual stream.       |

## getXlsxStreams
Async [generator](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Statements/function*) which yields transform streams of the sheets.

### Options

| option   | type                   | description                                 |
|----------|------------------------|---------------------------------------------|
| filePath | `string`               | Path to the XLSX file                       |
| sheets   | array of sheet objects | Options of sheet object can be found below. |

#### Sheet object

| option          | type                 | description                                                                                                                                                                                                                                                               |
|-----------------|----------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| id              | `string` or `number` | If `string` is passed, finds sheet by it's name. If `number`, finds sheet by it's index.                                                                                                                                                                                  |
| withHeader      | `boolean`            | If `true`, column names will be taken from the first sheet row. If duplicated header name is found, column name will be prepended with column letter to maintain uniqueness. 0-based row location can be passed to this option if header is not located on the first row. |
| ignoreEmpty     | `boolean`            | If `true`, empty rows won't be emitted.                                                                                                                                                                                                                                   |
| fillMergedCells | `boolean`            | If `true`, merged cells will have the same value (by default, only the first cell of merged cells is filled with value). **Warning!** Enabling this feature may increase streaming time because file must be processed to detect merged cells before actual stream.       |

## getWorksheets
Returns array of sheets with `name` and `hidden` info.

### Options

| option   | type     | description           |
|----------|----------|-----------------------|
| filePath | `string` | Path to the XLSX file |

## Building

You can build `js` source by using `npm run build` command.

## Testing

Tests can be run by using `npm test` command.
