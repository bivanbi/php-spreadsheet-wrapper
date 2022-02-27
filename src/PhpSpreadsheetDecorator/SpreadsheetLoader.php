<?php

namespace KignOrg\PhpSpreadsheetDecorator;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;

class SpreadsheetLoader
{
    /**
     * @throws Exception
     */
    public static function loadFromFile(string $filename, bool $readDataOnly = false): SpreadsheetDecorator
    {
        $inputFileType = IOFactory::identify($filename);
        $reader = IOFactory::createReader($inputFileType);
        $reader->setReadDataOnly($readDataOnly);
        return new SpreadsheetDecoratorImpl($reader->load($filename));
    }
}
