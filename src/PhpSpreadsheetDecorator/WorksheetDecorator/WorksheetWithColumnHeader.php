<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use PhpOffice\PhpSpreadsheet\Cell\Cell;

interface WorksheetWithColumnHeader
{
    public function setHeaderRow(int $row): WorksheetWithColumnHeader;

    public function setFirstDataRow(int $row): WorksheetWithColumnHeader;

    public function setLastDataRow(int $row): WorksheetWithColumnHeader;

    public function getColumnMap(): array;

    public function getRowIteratorWithColumnName($startRow = 1, $endRow = null): RowIteratorWithColumnName;

    public function getCellByColumnNameAndRow(string $column, int $row): Cell;
}
