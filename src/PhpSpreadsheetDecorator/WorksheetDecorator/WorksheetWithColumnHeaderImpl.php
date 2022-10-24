<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class WorksheetWithColumnHeaderImpl implements WorksheetWithColumnHeader
{
    protected Worksheet $worksheet;
    protected int $headerRow = 1;
    protected int $firstDataRow = 2;

    protected ?array $columnMap = null;
    protected ?array $columnNameToIndexMap = null;

    public function __construct(Worksheet $worksheet)
    {
        $this->worksheet = $worksheet;
    }

    public function getWorksheet(): Worksheet
    {
        return $this->worksheet;
    }

    public function setHeaderRow(int $row): WorksheetWithColumnHeader
    {
        $this->headerRow = $row;
        $this->invalidateCache();
        return $this;
    }

    public function setFirstDataRow(int $row): WorksheetWithColumnHeader
    {
        $this->firstDataRow = $row;
        return $this;
    }

    public function getColumnMap(): array
    {
        if (is_null($this->columnMap)) {
            $this->initColumnMap();
        }
        return $this->columnMap;
    }

    public function getColumnNameByAddress(string $columnAddress): string
    {
        $this->exceptOnInvalidColumnAddress($columnAddress);
        return array_search($columnAddress, $this->getColumnMap());
    }

    public function getRowIterator($startRow = null, $endRow = null): RowIteratorWithColumnName
    {
        return new RowIteratorWithColumnNameImpl($this, ($startRow ?? $this->firstDataRow), $endRow);
    }

    public function getCellByColumnNameAndRow(string $columnName, int $rowIndex): Cell
    {
        $this->exceptOnColumnNameNotFound($columnName);
        return $this->worksheet->getCell($this->getColumnAddressByName($columnName) . $rowIndex);
    }

    public function getColumnAddressByName(string $columnName): string
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

    protected function exceptOnInvalidColumnAddress(string $columnAddress): void
    {
        if (!$this->isValidColumnAddress($columnAddress)) {
            throw new InvalidArgumentException('No column name found for index');
        }
    }

    protected function exceptOnExistingColumnName(string $columnName): void
    {
        if (array_key_exists($columnName, ($this->columnMap ?? []))) {
            throw new InvalidArgumentException('Duplicate column name');
        }
    }

    public function isValidColumnAddress(string $columnAddress): bool
    {
        return in_array($columnAddress, $this->getColumnMap()) !== false;
    }

    /**
     * @throws Exception
     */
    public function addColumn(string $columnName, string $columnAddress = null): void
    {
        $columnAddress = $this->getFirstColumnAddressWithoutHeader();
        $this->worksheet->setCellValue($columnAddress . $this->headerRow, $columnName);
        $this->invalidateCache();
    }

    /**
     * @throws Exception
     */
    public function updateRow(int $rowIndex, array $values): void
    {
        foreach ($values as $columnName => $value) {
            $this->setCellValue($columnName, $rowIndex, $value);
        }
    }

    /**
     * @throws Exception
     */
    public function setCellValue(string $columnName, int $rowIndex, mixed $value = null): void
    {
        $this->exceptOnColumnNameNotFound($columnName);
        $map = $this->getColumnNameToIndexMap();
        $this->getWorksheet()->setCellValueByColumnAndRow($map[$columnName], $rowIndex, $value);
    }

    /**
     * @throws Exception
     */
    protected function getFirstColumnAddressWithoutHeader(): string
    {
        $columnIndex = $this->getFirstColumnIndexWithoutHeader();
        return Coordinate::stringFromColumnIndex($columnIndex);
    }

    /**
     * @throws Exception
     */
    protected function getFirstColumnIndexWithoutHeader(): string
    {
        return $this->getFirstMissingIndex($this->getColumnNameToIndexMap());
    }

    /**
     * @throws Exception
     */
    protected function getColumnNameToIndexMap(): array
    {
        if (is_null($this->columnNameToIndexMap)) {
            $this->initColumnNameToIndexMap();
        }
        return $this->columnNameToIndexMap;
    }

    /**
     * @throws Exception
     */
    public function initColumnNameToIndexMap(): void
    {
        $this->columnNameToIndexMap = [];
        foreach ($this->getColumnMap() as $name => $address) {
            $this->columnNameToIndexMap[$name] = Coordinate::columnIndexFromString($address);
        }
    }

    protected function getFirstMissingIndex(array $indices): int
    {
        sort($indices);
        /** @noinspection PhpStatementHasEmptyBodyInspection */
        for ($i = 1; in_array($i, $indices); $i++) ;
        return $i;
    }

    protected function invalidateCache(): void
    {
        $this->columnMap = null;
        $this->columnNameToIndexMap = null;
    }
}
