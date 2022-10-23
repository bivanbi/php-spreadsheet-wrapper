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
        $worksheet = $this->getColumnHeaderWorksheetWithInterleave();
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

    public function testGetColumnNameByAddress_withInvalidAddress()
    {
        $this->expectException(InvalidArgumentException::class);
        $this->worksheet->getColumnNameByAddress('Z');
    }

    public function testGetColumnAddressByName_withInvalidName()
    {
        $this->expectException(InvalidArgumentException::class);
        $this->worksheet->getColumnAddressByName("NON-existing-column-name");
    }

    public function testGetColumnAddressByName_withValidName()
    {
        foreach (TestConstants::HEADER_COLUMNS as $ColumnAddress => $columnName) {
            self::assertEquals($ColumnAddress, $this->worksheet->getColumnAddressByName($columnName));
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
        $expectedColumnMap = array_merge(TestConstants::HEADER_COLUMNS, ['E' => 'New Column']);
        $this->worksheet->addColumn($expectedColumnMap['E']);
        $actualColumnMap = $this->worksheet->getColumnMap();
        self::assertEquals(array_flip($expectedColumnMap), $actualColumnMap);
    }

    public function testAddColumn_viaAddColumnMethod_withInterleavedHeader()
    {
        $worksheet = $this->getColumnHeaderWorksheetWithInterleave();
        $expectedColumnMap = array_merge(TestConstants::HEADER_COLUMNS, ['B' => 'New Column']);
        $worksheet->addColumn($expectedColumnMap['B']);
        $actualColumnMap = $worksheet->getColumnMap();
        self::assertEquals(array_flip($expectedColumnMap), $actualColumnMap);
    }

    private function getColumnHeaderWorksheetWithInterleave(): WorksheetWithColumnHeader
    {
        $worksheet = $this->spreadsheet->getWorksheetWithColumnHeader(TestConstants::COLUMN_HEADER_WITH_INTERLEAVE_TEST_WORKSHEET_NAME);
        $worksheet->setHeaderRow(TestConstants::HEADER_ROW);
        $worksheet->setFirstDataRow(TestConstants::FIRST_DATA_ROW);
        return $worksheet;
    }
}
