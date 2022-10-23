<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class WorksheetWithColumnHeaderImpl extends Worksheet implements WorksheetWithColumnHeader
{
    protected Worksheet $worksheet;
    protected int $headerRow = 1;
    protected int $firstDataRow = 2;

    protected array $columnMap = [];

    public function __construct(Worksheet $worksheet)
    {
        parent::__construct($worksheet->getParent(), $worksheet->getTitle());
        $this->worksheet = $worksheet;
    }

    public function getWorksheet(): Worksheet
    {
        return $this->worksheet;
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


    public function getColumnMap(): array
    {
        if (sizeof($this->columnMap) === 0) {
            $this->initColumnMap();
        }
        return $this->columnMap;
    }

    public function getColumnName(string $columnIndex): string
    {
        $this->exceptOnInvalidColumnIndex($columnIndex);
        return array_search($columnIndex, $this->getColumnMap());
    }

    public function getRowIterator($startRow = null, $endRow = null): RowIteratorWithColumnName
    {
        return new RowIteratorWithColumnNameImpl($this, ($startRow ?? $this->firstDataRow), $endRow);
    }

    public function getCellByColumnNameAndRow(string $columnName, int $rowIndex): Cell
    {
        $this->exceptOnColumnNameNotFound($columnName);
        return $this->worksheet->getCell($this->getColumnIndex($columnName) . $rowIndex);
    }

    public function getColumnIndex(string $columnName): string
    {
        $this->exceptOnColumnNameNotFound($columnName);
        return $this->getColumnMap()[$columnName];
    }

    protected function initColumnMap()
    {
        $this->columnMap = [];
        $rowIterator = $this->worksheet->getRowIterator($this->headerRow);
        $cellIterator = $rowIterator->current()->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(true);

        foreach ($cellIterator as $index => $cell) {
            $name = $cell->getValue();
            $this->exceptOnExistingColumnName($name);
            $this->columnMap[$name] = $index;
        }
    }

    protected function exceptOnColumnNameNotFound(string $columnName): void
    {
        if (!array_key_exists($columnName, $this->getColumnMap())) {
            throw new InvalidArgumentException('Column name not found');
        }
    }

    protected function exceptOnInvalidColumnIndex(string $columnIndex): void
    {
        if (!$this->isValidColumnIndex($columnIndex)) {
            throw new InvalidArgumentException('No column name found for index');
        }
    }

    protected function exceptOnExistingColumnName(string $columnName): void
    {
        if (array_key_exists($columnName, ($this->columnMap ?? []))) {
            throw new InvalidArgumentException('Duplicate column name');
        }
    }

    public function isValidColumnIndex(string $columnIndex): bool
    {
        return in_array($columnIndex, $this->getColumnMap()) !== false;
    }
}
