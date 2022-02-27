<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class WorksheetWithColumnHeaderImpl extends Worksheet implements WorksheetWithColumnHeader
{
    public function __construct(Worksheet $worksheet)
    {
        parent::__construct($worksheet->getParent(), $worksheet->getTitle());
    }

    public function getRowIteratorWithColumnName($startRow = 1, $endRow = null): RowIteratorWithColumnName
    {
        // TODO: Implement getRowIteratorWithColumnName() method.
        return new RowIteratorWithColumnNameImpl();
    }

    public function getCellByColumnNameAndRow(string $column, int $row): Cell
    {
        // TODO: Implement getCellByColumnNameAndRow() method.
        $columnId = "A";
        return parent::getCellByColumnAndRow(1, $row);
    }
}
