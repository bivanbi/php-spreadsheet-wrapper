<?php

namespace KignOrg\PhpSpreadsheetDecorator;

use KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator\WorksheetWithColumnHeader;

interface SpreadsheetDecorator
{
    public function getWorksheetWithColumnHeader(string $name): WorksheetWithColumnHeader;
}
