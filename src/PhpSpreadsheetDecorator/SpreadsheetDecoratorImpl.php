<?php

namespace KignOrg\PhpSpreadsheetDecorator;

use KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator\WorksheetWithColumnHeader;
use KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator\WorksheetWithColumnHeaderImpl;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SpreadsheetDecoratorImpl extends Spreadsheet implements SpreadsheetDecorator
{
    protected Spreadsheet $spreadsheet;

    public function __construct(Spreadsheet $spreadsheet)
    {
        parent::__construct();
        $this->spreadsheet = $spreadsheet;
    }

    public function getWorksheetWithColumnHeader(string $name): WorksheetWithColumnHeader
    {
        return new WorksheetWithColumnHeaderImpl($this->spreadsheet->getSheetByName($name));
    }
}
