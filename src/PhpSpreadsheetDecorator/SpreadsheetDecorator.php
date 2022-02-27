<?php

namespace KignOrg\PhpSpreadsheetDecorator;

use KignOrg\PhpSpreadSheetDecorator\WorksheetDecorator\WorksheetWithColumnHeader;

interface SpreadsheetDecorator
{
    public function getWorksheetWithColumnHeader(string $name): WorksheetWithColumnHeader;
}
