# Php Spreadsheet Wrapper

A wrapper around
[https://phpoffice.github.io/PhpSpreadsheet](https://phpoffice.github.io/PhpSpreadsheet)
library, with focus on data input and output: reading from and writing to worksheets
utilizing column / row header.

## Main functions

- Map column header to column letter
- Read sheet data into associative array ('dictionary') where column header is the key
- Populate sheet with from associative array in the same manner

## What Php Spreadsheet Wrapper is not

This wrapper only implements functions necessary to realize the above functions.
It deliberately does not, for instance, expose or augment formatting functions.
If needed, one can always access the contained PhpSpreadsheet objects and carry
out such tasks as would normally do with 'bare' PhpSpreadsheet.
