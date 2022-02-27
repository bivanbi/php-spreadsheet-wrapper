<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use KignOrg\PhpSpreadsheetDecorator\SpreadsheetLoader;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PHPUnit\Framework\TestCase;

class WorksheetWithColumnHeaderTest extends TestCase
{
    const XLS_FILENAME = 'tests/testsheet.xlsx';
    const TEST_WORKSHEET_NAME = 'TestSheet';
    const HEADER_ROW = 1;
    const HEADER_COLUMN = "A";

    /**
     * @throws Exception
     */
    protected function setUp(): void
    {
        print_r(__DIR__);
        print_r(getcwd());
        print_r(new TestClass());
        $this->spreadsheetDecorator = SpreadsheetLoader::loadFromFile(self::XLS_FILENAME);
    }

    public function testName()
    {
        $this->assertTrue(true, "yeah");
    }

}
