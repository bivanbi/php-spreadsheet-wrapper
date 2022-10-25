<?php

namespace KignOrg\PhpSpreadsheetDecorator;

use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Reader\Exception as ReaderException;
use PHPUnit\Framework\TestCase;

class SpreadsheetDecoratorTest extends TestCase
{
    protected SpreadsheetDecorator $spreadsheet;
    protected string $tempFileName;

    /**
     * @throws ReaderException
     */
    protected function setUp(): void
    {
        $classShortName = substr(get_called_class(), strrpos(get_called_class(), '\\') + 1);
        $this->spreadsheet = SpreadsheetLoader::loadFromFile(TestConstants::TEST_XLS_FILENAME);
        $this->tempFileName = tempnam("/tmp", $classShortName);
    }

    protected function tearDown(): void
    {
        parent::tearDown();
        unlink($this->tempFileName);
    }


    public function testGetWorksheetWithColumnHeader_withInvalidSheetName()
    {
        $this->expectException(InvalidArgumentException::class);
        $this->spreadsheet->getWorksheetWithColumnHeader('absolutely invalid sheet name');
    }

    /**
     * @throws ReaderException
     */
    public function testSaveToFile()
    {
        $worksheet = $this->spreadsheet->getWorksheetWithColumnHeader(TestConstants::COLUMN_HEADER_WITH_NO_DATA_ROWS_TEST_WORKSHEET_NAME);
        $worksheet->setHeaderRow(TestConstants::HEADER_ROW);
        $worksheet->setFirstDataRow(TestConstants::FIRST_DATA_ROW);
        foreach (TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT as $rowIndex => $rowData) {
            $worksheet->updateRow($rowIndex, $rowData);
        }
        $this->spreadsheet->saveToFile($this->tempFileName);

        $worksheet = SpreadsheetLoader::loadFromFile($this->tempFileName)->getWorksheetWithColumnHeader(TestConstants::COLUMN_HEADER_WITH_NO_DATA_ROWS_TEST_WORKSHEET_NAME);
        $worksheet->setHeaderRow(TestConstants::HEADER_ROW);
        $worksheet->setFirstDataRow(TestConstants::FIRST_DATA_ROW);

        $rowIterator = $worksheet->getRowIterator();
        foreach ($rowIterator as $row) {
            $expected = TestConstants::COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT[$row->getRowIndex()];
            $actual = $row->asArray();
            self::assertEquals($expected, $actual);
        }
    }
}
