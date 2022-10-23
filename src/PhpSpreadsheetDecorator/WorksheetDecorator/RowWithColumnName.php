<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

interface RowWithColumnName
{
    public function asArray(): array;

    public function getRowIndex(): int;

    public function getCellIterator($startColumn = null, $endColumn = null): RowCellIterator;

    public function getWorksheet(): Worksheet;

    public function getWorksheetWithColumnHeader(): WorksheetWithColumnHeader;

    public function getColumnMap(): array;
}
