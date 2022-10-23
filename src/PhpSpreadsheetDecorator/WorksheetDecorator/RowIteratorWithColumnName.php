<?php

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use Iterator;
use PhpOffice\PhpSpreadsheet\Worksheet\Row;

/**
 * @implements Iterator<int, Row>
 */
interface RowIteratorWithColumnName extends Iterator
{
    public function current(): RowWithColumnName;
}
