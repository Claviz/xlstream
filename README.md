# xlstream

Memory-efficiently turns XLSX file into a [transform stream](https://nodejs.org/api/stream.html#stream_duplex_and_transform_streams) with all it's benefits.

* Stream is **pausable**.
* Emits all **default events** (`data`, `end`, etc.)
* Returns **raw** and **formatted** row data in just one `data` event.

## Installation
```
npm install xlstream
```

## Example
`source.xlsx` contents:
| A     | B   |
| ----- | --- |
| hello | 123 |
Where `123` is a `123.123` number formatted to be rounded to integer.

Script:
```javascript
const getXlsxStream = 'xlstream';

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
    } 
}
```

## Options

| option   | type                 | description                                                                              |
| -------- | -------------------- | ---------------------------------------------------------------------------------------- |
| filePath | `string`             | Path to the XLSX file                                                                    |
| sheet    | `string` or `number` | If `string` is passed, finds sheet by it's name. If `number`, finds sheet by it's index. |

## Testing

Tests can be run by using `npm test` command.