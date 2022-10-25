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
}
