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

    public function getColumnIndex(string $columnName): string;

    public function getColumnName(string $columnIndex): string;

    public function getRowIterator($startRow = null, $endRow = null): RowIteratorWithColumnName;

    public function getCellByColumnNameAndRow(string $columnName, int $rowIndex): Cell;

    public function isValidColumnIndex(string $columnIndex): bool;
}
