<?php /** @noinspection PhpPureAttributeCanBeAddedInspection */

namespace KignOrg\PhpSpreadsheetDecorator\WorksheetDecorator;

use Iterator;
use PhpOffice\PhpSpreadsheet\Worksheet\RowIterator;

/**
 * @implements Iterator<int, RowWithColumnName>
 */
class RowIteratorWithColumnNameImpl implements RowIteratorWithColumnName
{
    private RowIterator $rowIterator;
    private WorksheetWithColumnHeader $worksheetWithColumnHeader;

    public function __construct(WorksheetWithColumnHeader $worksheet, $startRow = 1, $endRow = null)
    {
        $this->rowIterator = new RowIterator($worksheet->getWorksheet(), $startRow, $endRow);
        $this->worksheetWithColumnHeader = $worksheet;
    }

    public function current(): RowWithColumnName
    {
        return new RowWithColumnNameImpl($this->worksheetWithColumnHeader, $this->rowIterator->current());
    }

    public function next()
    {
        $this->rowIterator->next();
    }

    public function key()
    {
        return $this->rowIterator->key();
    }

    public function valid(): bool
    {
        return $this->rowIterator->valid();
    }

    public function rewind()
    {
        $this->rowIterator->rewind();
    }
}
