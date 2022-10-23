<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use InvalidArgumentException;
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

    const EXPECTED_VALUES = [
        2 => ["Name" => "Barbara", "Age" => 19, "Home Country" => "Germany", "Occupation" => "University Student"],
        3 => ["Name" => "Emmylou", "Age" => 24, "Home Country" => "Netherlands", "Occupation" => "Writer"],
        4 => ["Name" => "Johnny", "Age" => 65, "Home Country" => "USA", "Occupation" => "Musician"],
        5 => ["Name" => "Sammy", "Age" => 34, "Home Country" => "Canada", "Occupation" => "Private Entrepreneur"],
        6 => ["Name" => "Walt", "Age" => 102, "Home Country" => "Ireland", "Occupation" => "Retired"],
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
        $this->worksheet->setHeaderRow(self::HEADER_ROW);
        $this->worksheet->setFirstDataRow(self::FIRST_DATA_ROW);
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
        $this->assertEquals(array_flip(self::COLUMNS), $actual);
    }

    public function testRowIterator()
    {
        $rowIterator = $this->worksheet->getRowIterator();
        foreach ($rowIterator as $row) {
            $expected = self::EXPECTED_VALUES[$row->getRowIndex()];
            $actual = $row->asArray();
            self::assertEquals($expected, $actual);
        }
    }

    public function testGetColumnName_withInvalidIndex()
    {
        $this->expectException(InvalidArgumentException::class);
        $this->worksheet->getColumnName('Z');
    }

    public function testGetColumnIndex_withInvalidName()
    {
        $this->expectException(InvalidArgumentException::class);
        $this->worksheet->getColumnIndex("NON-existing-column-name");
    }

    public function testGetColumnIndex_withValidName()
    {
        foreach (self::COLUMNS as $columnIndex => $columnName) {
            self::assertEquals($columnIndex, $this->worksheet->getColumnIndex($columnName));
        }
    }

    public function testGetCellByColumnNameAndRow()
    {
        foreach (self::EXPECTED_VALUES as $rowIndex => $columns) {
            foreach ($columns as $columnName => $expectedValue) {
                $cell = $this->worksheet->getCellByColumnNameAndRow($columnName, $rowIndex);
                self::assertEquals($expectedValue, $cell->getValue());
            }
        }
    }
}
