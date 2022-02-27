<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use PhpOffice\PhpSpreadsheet\Cell\Cell;

interface WorksheetWithColumnHeader
{
    public function getRowIteratorWithColumnName($startRow = 1, $endRow = null): RowIteratorWithColumnName;

    public function getCellByColumnNameAndRow(string $column, int $row): Cell;
}
