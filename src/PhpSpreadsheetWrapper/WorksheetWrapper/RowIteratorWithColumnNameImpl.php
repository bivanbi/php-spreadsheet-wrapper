<?php /** @noinspection PhpPureAttributeCanBeAddedInspection */

namespace KignOrg\PhpSpreadsheetWrapper\WorksheetWrapper;

use Iterator;
use PhpOffice\PhpSpreadsheet\Worksheet\RowIterator;
use ReturnTypeWillChange;

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

    #[ReturnTypeWillChange] public function next()
    {
        $this->rowIterator->next();
    }

    public function key(): int
    {
        return $this->rowIterator->key();
    }

    public function valid(): bool
    {
        return $this->rowIterator->valid();
    }

    #[ReturnTypeWillChange] public function rewind()
    {
        $this->rowIterator->rewind();
    }
}
