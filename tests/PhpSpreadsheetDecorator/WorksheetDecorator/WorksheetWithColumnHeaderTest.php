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

    /**
     * @throws ReaderException
     */
    protected function setUp(): void
    {
        $this->spreadsheet = SpreadsheetLoader::loadFromFile(TestConstants::TEST_XLS_FILENAME);
    }

    public function testGetColumnMap_withHeaderRowBeyondHighest()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $this->expectException(SpreadsheetException::class);
        $worksheet->setHeaderRow(10000);
        $worksheet->getColumnMap();
    }

    public function testGetColumnMap_withValidHeaderRow()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $actual = $worksheet->getColumnMap();
        $this->assertEquals(array_flip(TestConstants::HEADER_COLUMNS), $actual);
    }

    public function testGetColumnMap_withInterleavedHeader()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_WITH_INTERLEAVE_TEST_WORKSHEET_NAME);
        $actual = $worksheet->getColumnMap();
        $this->assertEquals(array_flip(TestConstants::HEADER_COLUMNS_WITH_INTERLEAVE), $actual);
    }

    public function testRowIterator()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $rowIterator = $worksheet->getRowIterator();
        foreach ($rowIterator as $row) {
            $expected = TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT[$row->getRowIndex()];
            $actual = $row->asArray();
            self::assertEquals($expected, $actual);
        }
    }

    public function testGetColumnNameByAddress_withInvalidAddress()
    {
        $this->expectException(InvalidArgumentException::class);
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $worksheet->getColumnNameByAddress('Z');
    }

    public function testGetColumnAddressByName_withInvalidName()
    {
        $this->expectException(InvalidArgumentException::class);
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $worksheet->getColumnAddressByName("NON-existing-column-name");
    }

    public function testGetColumnAddressByName_withValidName()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        foreach (TestConstants::HEADER_COLUMNS as $ColumnAddress => $columnName) {
            self::assertEquals($ColumnAddress, $worksheet->getColumnAddressByName($columnName));
        }
    }

    public function testGetCellByColumnNameAndRow()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        foreach (TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT as $rowIndex => $columns) {
            foreach ($columns as $columnName => $expectedValue) {
                $cell = $worksheet->getCellByColumnNameAndRow($columnName, $rowIndex);
                self::assertEquals($expectedValue, $cell->getValue());
            }
        }
    }

    public function testAddColumn()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $expectedColumnMap = array_merge(TestConstants::HEADER_COLUMNS, ['E' => 'New Column']);
        $worksheet->addColumn($expectedColumnMap['E']);
        $actualColumnMap = $worksheet->getColumnMap();
        self::assertEquals(array_flip($expectedColumnMap), $actualColumnMap);
    }

    public function testAddColumn_withDuplicateColumnName()
    {
        $this->expectException(InvalidArgumentException::class);
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_TEST_WORKSHEET_NAME);
        $existingColumnName = current(array_keys(current(TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT)));
        $worksheet->addColumn($existingColumnName);
    }

    public function testAddColumn_withInterleavedHeader()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_WITH_INTERLEAVE_TEST_WORKSHEET_NAME);
        $expectedColumnMap = array_merge(TestConstants::HEADER_COLUMNS, ['B' => 'New Column']);
        $worksheet->addColumn($expectedColumnMap['B']);
        $actualColumnMap = $worksheet->getColumnMap();
        self::assertEquals(array_flip($expectedColumnMap), $actualColumnMap);
    }

    public function testUpdateRow_withInvalidColumnName()
    {
        $this->expectException(InvalidArgumentException::class);
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_WITH_NO_DATA_ROWS_TEST_WORKSHEET_NAME);
        $rowIndex = TestConstants::FIRST_DATA_ROW;
        $worksheet->updateRow($rowIndex, ['invalid header' => 'some data']);
    }

    public function testUpdateRow_withValidData()
    {
        $worksheet = $this->getWorksheet(TestConstants::COLUMN_HEADER_WITH_NO_DATA_ROWS_TEST_WORKSHEET_NAME);
        foreach (TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT as $rowIndex => $rowData) {
            $worksheet->updateRow($rowIndex, $rowData);
        }

        $rowIterator = $worksheet->getRowIterator();
        foreach ($rowIterator as $row) {
            $expected = TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT[$row->getRowIndex()];
            $actual = $row->asArray();
            self::assertEquals($expected, $actual);
        }
    }

    private function getWorksheet(string $worksheetName): WorksheetWithColumnHeader
    {
        $worksheet = $this->spreadsheet->getWorksheetWithColumnHeader($worksheetName);
        $worksheet->setHeaderRow(TestConstants::HEADER_ROW);
        $worksheet->setFirstDataRow(TestConstants::FIRST_DATA_ROW);
        return $worksheet;
    }
}
