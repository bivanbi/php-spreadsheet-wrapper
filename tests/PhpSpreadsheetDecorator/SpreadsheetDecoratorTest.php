<?php

namespace KignOrg\PhpSpreadsheetDecorator;

use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Reader\Exception as ReaderException;
use PHPUnit\Framework\TestCase;

class SpreadsheetDecoratorTest extends TestCase
{
    protected SpreadsheetDecorator $spreadsheet;

    /**
     * @throws ReaderException
     */
    protected function setUp(): void
    {
        $this->spreadsheet = SpreadsheetLoader::loadFromFile(TestConstants::TEST_XLS_FILENAME);
    }

    public function testGetWorksheetWithColumnHeader_withInvalidSheetName()
    {
        $this->expectException(InvalidArgumentException::class);
        $this->spreadsheet->getWorksheetWithColumnHeader('absolutely invalid sheet name');
    }
}
