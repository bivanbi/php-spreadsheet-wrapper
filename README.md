# Php Spreadsheet Wrapper

A wrapper around
[https://phpoffice.github.io/PhpSpreadsheet](https://phpoffice.github.io/PhpSpreadsheet)
library, with focus on data input and output: reading from and writing to worksheets
utilizing column / row header.

## Main functions

- Map column header to column letter
- Read sheet data into associative array ('dictionary') where column header is the key
- Populate sheet with data from associative array in the same manner

## What Php Spreadsheet Wrapper is not

This wrapper only implements functions necessary to realize the above functions.
It deliberately does not, for instance, expose or augment formatting functions.
If needed, one can always access the contained PhpSpreadsheet objects and interact
with them directly.

## Installation

Add Php Spreadsheet Wrapper as a requirement to the project composer.json like this:

```json
{
  "repositories": [
    {
      "url": "https://github.com/bivanbi/php-spreadsheet-wrapper",
      "type": "git"
    }
  ],
  "require": {
    "kign.org/php-spreadsheet-wrapper": "master"
  }
}
```

_Note: it is not recommended to use "master" version, as there can be breaking changes.
Consult GitHub for actual releases._

## How to Use

Check out [tests](tests) for working examples.

### Read from Worksheet with column headers

```php
use KignOrg\PhpSpreadsheetWrapper\SpreadsheetLoader;
use KignOrg\PhpSpreadsheetWrapper\SpreadsheetWrapper;

$spreadsheet = SpreadsheetLoader::loadFromFile("/path/to/spreadsheet.xlsx"); // Open file
$worksheet = $spreadsheet->getWorksheetWithColumnHeader("Name of My Sheet"); // Open worksheet to be read
// $worksheet->setHeaderRow(1);    // Optional, set header row. Default is 1 (first row).
// $worksheet->setFirstDataRow(2); // Optional, first data row. Default is 2 (second row).

$rowIterator = $worksheet->getRowIterator();
foreach ($rowIterator as $row) { // Iterate over rows. Recommended for large sheets.
    $data = $row->asArray(); // get row as associative array (dictionary) where column header is the key
    // do whatever is needed to do with the data
}

$individualCellValue = $worksheet->getCellByColumnNameAndRow('City', 3); // get a specific cell's value

```

### Write to Worksheet with column headers

```php
use KignOrg\PhpSpreadsheetWrapper\SpreadsheetLoader;
use KignOrg\PhpSpreadsheetWrapper\SpreadsheetWrapper;

$spreadsheet = SpreadsheetLoader::loadFromFile("/path/to/spreadsheet.xlsx"); // Open file
$worksheet = $spreadsheet->getWorksheetWithColumnHeader("Name of My Sheet"); // Open worksheet to be filled

// $data = ['2' => ['City' => 'London', 'Country' => 'England' ..], '2' => ['City => 'Luxembourg', 'Country' => 'Luxembourg'] ... ]
foreach ($data as $rowIndex => $rowData) { 
    $worksheet->updateRow($rowIndex, $rowData);
}

$this->spreadsheet->saveToFile($this->tempFileName); // Write spreadsheet to file
```
