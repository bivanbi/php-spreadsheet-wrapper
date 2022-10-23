<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

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
     * @param string $columnName
     * @param string|null $atIndex index of column explicitly. Defaults to first unused (column without header) if not set
     */
    public function addColumn(string $columnName, string $atIndex = null): void;

    //public function getFirstColumnIndexWithoutHeader(): string;
}
