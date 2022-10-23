<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use InvalidArgumentException;
use KignOrg\PhpSpreadsheetDecorator\SpreadsheetDecorator;
use KignOrg\PhpSpreadsheetDecorator\SpreadsheetLoader;
use KignOrg\PhpSpreadsheetDecorator\TestConstants;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Reader\Exception as ReaderException;
use PHPUnit\Framework\TestCase;

class WorksheetWithColumnHeaderTest extends TestCase
{
    protected SpreadsheetDecorator $spreadsheet;
    protected WorksheetWithColumnHeader $worksheet;

    /**
     * @throws ReaderException
     */
    protected function setUp(): void
    {
        $this->spreadsheet = SpreadsheetLoader::loadFromFile(TestConstants::TEST_XLS_FILENAME);
        $this->worksheet = $this->spreadsheet->getWorksheetWithColumnHeader(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $this->worksheet->setHeaderRow(TestConstants::HEADER_ROW);
        $this->worksheet->setFirstDataRow(TestConstants::FIRST_DATA_ROW);
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
        $this->assertEquals(array_flip(TestConstants::HEADER_COLUMNS), $actual);
    }

    public function testGetColumnMap_withInterleavedHeader()
    {
        $worksheet = $this->spreadsheet->getWorksheetWithColumnHeader(TestConstants::COLUMN_HEADER_WITH_INTERLEAVE_TEST_WORKSHEET_NAME);
        $worksheet->setHeaderRow(TestConstants::HEADER_ROW);
        $worksheet->setFirstDataRow(TestConstants::FIRST_DATA_ROW);
        $actual = $worksheet->getColumnMap();
        $this->assertEquals(array_flip(TestConstants::HEADER_COLUMNS_WITH_INTERLEAVE), $actual);
    }

    public function testRowIterator()
    {
        $rowIterator = $this->worksheet->getRowIterator();
        foreach ($rowIterator as $row) {
            $expected = TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT[$row->getRowIndex()];
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
        foreach (TestConstants::HEADER_COLUMNS as $columnIndex => $columnName) {
            self::assertEquals($columnIndex, $this->worksheet->getColumnIndex($columnName));
        }
    }

    public function testGetCellByColumnNameAndRow()
    {
        foreach (TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT as $rowIndex => $columns) {
            foreach ($columns as $columnName => $expectedValue) {
                $cell = $this->worksheet->getCellByColumnNameAndRow($columnName, $rowIndex);
                self::assertEquals($expectedValue, $cell->getValue());
            }
        }
    }

    public function testAddColumn_viaAddColumnMethod()
    {
        $this->worksheet->addColumn("New Column");
    }
}
