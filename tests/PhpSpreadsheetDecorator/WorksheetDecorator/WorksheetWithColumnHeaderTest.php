<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use KignOrg\PhpSpreadsheetDecorator\SpreadsheetDecorator;
use KignOrg\PhpSpreadsheetDecorator\SpreadsheetLoader;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Reader\Exception as ReaderException;
use PHPUnit\Framework\TestCase;

class WorksheetWithColumnHeaderTest extends TestCase
{
    const XLS_FILENAME = 'tests/testsheet.xlsx';
    const TEST_WORKSHEET_NAME = 'TestSheet';
    const HEADER_ROW = 1;
    const HEADER_COLUMN = "A";
    const FIRST_DATA_ROW = 2;
    const LAST_DATA_ROW = null;

    const COLUMNS = [
        "A" => "Name",
        "B" => "Age",
        "C" => "Home Country",
        "D" => "Occupation",
    ];

    protected SpreadsheetDecorator $spreadsheet;
    protected WorksheetWithColumnHeader $worksheet;

    /**
     * @throws ReaderException
     */
    protected function setUp(): void
    {
        $this->spreadsheet = SpreadsheetLoader::loadFromFile(self::XLS_FILENAME);
        $this->worksheet = $this->spreadsheet->getWorksheetWithColumnHeader(self::TEST_WORKSHEET_NAME);
        $this->worksheet
            ->setHeaderRow(self::HEADER_ROW)
            ->setFirstDataRow(self::FIRST_DATA_ROW)
            ->setLastDataRow(self::LAST_DATA_ROW);
    }

    public function testGetColumnMap_withHeaderRowBeyondHighest()
    {
        $this->expectException(SpreadsheetException::class);
        $this->worksheet->setHeaderRow(10000);
        $this->worksheet->getColumnMap();
    }

    public function testGetColumnMap_withValidHeaderRow()
    {
        $actual = $this->worksheet->getColumnMap();
        $this->assertEquals(self::COLUMNS, $actual);
    }
}
