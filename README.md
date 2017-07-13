## Introduction

Excel XLSX to CSV file converter. Node.js library and a command line utility. Lightweight.

## Getting Started

Install the package:
```bash
git clone https://github.com/papnkukn/xlsx2csv
npm install -g .
```

Convert the document:
```bash
xlsx2csv --verbose --sheet Sample path/to/sample.xlsx sample.csv
```

## Command Line

```
Usage:
  xlsx2csv [options] <input.xlsx> <output.csv>

Options:
  --force             Force overwrite file
  --data [type]       Data type to export: formula, value or display
  --sheet [name]      Sheet name to export
  --range [A1:C3]     Range to export
  --separator [char]  CSV column separator, e.g. '\t'
  --line-end [char]   End of line char(s), e.g. '\r\n'
  --help              Print this message
  --verbose           Enable detailed logging
  --version           Print version number

Examples:
  xlsx2csv --version
  xlsx2csv --verbose --range A1:M30 --separator , file.xlsx
  xlsx2csv --data formula --sheet Sample file.xlsx file.csv
```

## Using as library

```javascript
var xlsx2csv = require('node-xlsx2csv');
var options = { verbose: true, sheet: "Sample" };
xlsx2csv('path/to/sample.xlsx', options, function(error, result) {
  if (error) return console.error(error);
  console.log(result);
});
```

```
var options = {
  verbose: true|false
  data: type of data to export: "formula" to export cell formula, "value" to prefer cell value, or "display" for formatted value
  sheet: required, sheet name, e.g. "Sample"
  range: optional, cell range, e.g. "A1:M30"
  separator: CSV column separator, e.g. "," or ";" or "\t", default: ","
  lineEnd: end of line char(s), e.g. "\r\n" or "\r" or "\n", default: "\n"
};
```

<!--
## Example of output

```csv
xlsx2csv,,
1234,,
1,,"a, b, and c"
10,,"a \"quoted\" string"
99%,,
0.5625,,
```
-->