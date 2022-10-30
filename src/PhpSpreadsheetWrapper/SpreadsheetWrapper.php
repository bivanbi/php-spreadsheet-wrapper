<?php

namespace KignOrg\PhpSpreadsheetWrapper;

use KignOrg\PhpSpreadsheetWrapper\WorksheetWrapper\WorksheetWithColumnHeader;

interface SpreadsheetWrapper
{
    public function getWorksheetWithColumnHeader(string $name): WorksheetWithColumnHeader;

    public function saveToFile(string $filename): void;
}
