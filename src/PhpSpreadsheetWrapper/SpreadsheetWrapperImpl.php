<?php

namespace KignOrg\PhpSpreadsheetWrapper;

use InvalidArgumentException;
use KignOrg\PhpSpreadsheetWrapper\WorksheetWrapper\WorksheetWithColumnHeader;
use KignOrg\PhpSpreadsheetWrapper\WorksheetWrapper\WorksheetWithColumnHeaderImpl;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class SpreadsheetWrapperImpl implements SpreadsheetWrapper
{
    protected Spreadsheet $spreadsheet;

    public function __construct(Spreadsheet $spreadsheet)
    {
        $this->spreadsheet = $spreadsheet;
    }

    public function getWorksheetWithColumnHeader(string $name): WorksheetWithColumnHeader
    {
        $worksheet = $this->spreadsheet->getSheetByName($name);
        if ($worksheet == null) {
            throw new InvalidArgumentException("Failed to get sheet by name");
        }
        return new WorksheetWithColumnHeaderImpl($this->spreadsheet->getSheetByName($name));
    }

    /**
     * @throws Exception
     */
    public function saveToFile(string $filename): void
    {
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($filename);
    }
}
