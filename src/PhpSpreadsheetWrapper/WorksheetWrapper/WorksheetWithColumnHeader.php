<?php

namespace KignOrg\PhpSpreadsheetWrapper\WorksheetWrapper;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

interface WorksheetWithColumnHeader
{
    public function setHeaderRow(int $row): WorksheetWithColumnHeader;

    public function setFirstDataRow(int $row): WorksheetWithColumnHeader;

    public function getWorksheet(): Worksheet;

    public function getColumnMap(): array;

    public function getColumnAddressByName(string $columnName): string;

    public function getColumnNameByAddress(string $columnAddress): string;

    public function getRowIterator($startRow = null, $endRow = null): RowIteratorWithColumnName;

    public function getCellByColumnNameAndRow(string $columnName, int $rowIndex): Cell;

    public function isValidColumnAddress(string $columnAddress): bool;

    /**
     * Update multiple cells in a row at once.
     * Only columns in $values array will be updated.
     *
     * @param int $rowIndex the index of the row to be updated
     * @param array $values array of column name - values
     */
    public function updateRow(int $rowIndex, array $values): void;

    public function setCellValue(string $columnName, int $rowIndex, mixed $value = null): void;

    /**
     * Add a new column with column name. This will be written to the cell specified by
     * columnAddress + header row. Parameter columnAddress defaults to the first column
     * address that has an empty cell in header row.
     *
     * @param string $columnName
     * @param string|null $columnAddress
     */
    public function addColumn(string $columnName, string $columnAddress = null): void;

    /**
     * Fill worksheet with data array of associative arrays (dictionaries, 'rows'), beginning at $startRow.
     * If $startRow is omitted, if defaults to 'first data row'. The 'row' index in the input data
     * is not used, but the order is maintained. If sheet contains data in the given row range, it will
     * be overwritten, but only the columns that are defined in the actual input data.
     *
     * @param array $data
     * @param int|null $startRow
     */
    public function fillWorksheet(array $data, int $startRow = null): void;
}
