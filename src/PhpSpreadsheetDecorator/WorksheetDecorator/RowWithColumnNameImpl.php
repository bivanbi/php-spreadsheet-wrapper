<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use PhpOffice\PhpSpreadsheet\Worksheet\Row;
use PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class RowWithColumnNameImpl implements RowWithColumnName
{
    protected WorksheetWithColumnHeader $worksheetWithColumnHeader;
    protected Row $row;

    public function __construct(WorksheetWithColumnHeader $worksheet, Row $row)
    {
        $this->row = $row;
        $this->worksheetWithColumnHeader = $worksheet;
    }

    public function asArray(): array
    {
        $rowArray = [];
        $cellIterator = $this->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(true);

        foreach ($cellIterator as $columnIndex => $cell) {
            if ($this->worksheetWithColumnHeader->isValidColumnIndex($columnIndex)) {
                $columnName = $this->worksheetWithColumnHeader->getColumnName($columnIndex);
                $rowArray[$columnName] = $cell->getValue();
            }
        }
        return $rowArray;
    }

    public function getRowIndex(): int
    {
        return $this->row->getRowIndex();
    }

    public function getCellIterator($startColumn = null, $endColumn = null): RowCellIterator
    {
        $columnMap = array_values($this->getColumnMap());
        return $this->row->getCellIterator($startColumn ?? reset($columnMap), $endColumn ?? end($columnMap));
    }

    public function getWorksheet(): Worksheet
    {
        return $this->worksheetWithColumnHeader->getWorksheet();
    }

    public function getWorksheetWithColumnHeader(): WorksheetWithColumnHeader
    {
        return $this->worksheetWithColumnHeader;
    }

    public function getColumnMap(): array
    {
        return $this->worksheetWithColumnHeader->getColumnMap();
    }
}
