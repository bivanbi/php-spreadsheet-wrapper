<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class WorksheetWithColumnHeaderImpl extends Worksheet implements WorksheetWithColumnHeader
{
    protected Worksheet $worksheet;
    protected int $headerRow = 1;
    protected int $firstDataRow = 2;
    protected ?int $lastDataRow = null;

    protected array $columnMap = [];

    public function __construct(Worksheet $worksheet)
    {
        parent::__construct($worksheet->getParent(), $worksheet->getTitle());
        $this->worksheet = $worksheet;
    }

    public function setHeaderRow(int $row): WorksheetWithColumnHeader
    {
        $this->headerRow = $row;
        $this->columnMap = [];
        return $this;
    }

    public function setFirstDataRow(int $row): WorksheetWithColumnHeader
    {
        $this->firstDataRow = $row;
        return $this;
    }

    public function setLastDataRow(?int $row): WorksheetWithColumnHeader
    {
        $this->lastDataRow = $row;
        return $this;
    }

    public function getColumnMap(): array
    {
        if (sizeof($this->columnMap) === 0) {
            $this->initColumnMap();
        }
        return $this->columnMap;
    }

    protected function initColumnMap()
    {
        $this->columnMap = [];
        $rowIterator = $this->worksheet->getRowIterator($this->headerRow);
        $cellIterator = $rowIterator->current()->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(true);

        foreach ($cellIterator as $index => $cell) {
            $name = $cell->getValue();
            $this->columnMap[$index] = $name;
        }
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
