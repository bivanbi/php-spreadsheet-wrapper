<?php

namespace KignOrg\PhpSpreadsheetDecorator;

class TestConstants
{
    const TEST_XLS_FILENAME = 'tests/testsheet.xlsx';
    const COLUMN_HEADER_TEST_WORKSHEET_NAME = 'ColumnHeader';
    const COLUMN_HEADER_WITH_INTERLEAVE_TEST_WORKSHEET_NAME = 'ColumnHeaderWithInterleave';
    const COLUMN_HEADER_WITH_NO_DATA_ROWS_TEST_WORKSHEET_NAME = 'ColumnHeaderWithNoData';

    const HEADER_ROW = 1;
    const HEADER_COLUMN = 'A';
    const FIRST_DATA_ROW = 2;
    const LAST_DATA_ROW = null;

    const HEADER_COLUMNS = [
        'A' => 'Name',
        'B' => 'Age',
        'C' => 'Home Country',
        'D' => 'Occupation',
    ];

    const HEADER_COLUMNS_WITH_INTERLEAVE = [
        'A' => 'Name',
        'C' => 'Home Country',
        'D' => 'Occupation',
    ];

    const COLUMN_HEADER_TEST_SHEET_EXPECTED_CONTENT = [
        2 => ['Name' => 'Barbara', 'Age' => 19, 'Home Country' => 'Germany', 'Occupation' => 'University Student'],
        3 => ['Name' => 'Emmylou', 'Age' => 24, 'Home Country' => 'Netherlands', 'Occupation' => 'Writer'],
        4 => ['Name' => 'Johnny', 'Age' => 65, 'Home Country' => 'USA', 'Occupation' => 'Musician'],
        5 => ['Name' => 'Sammy', 'Age' => 34, 'Home Country' => 'Canada', 'Occupation' => 'Private Entrepreneur'],
        6 => ['Name' => 'Walt', 'Age' => 102, 'Home Country' => 'Ireland', 'Occupation' => 'Retired'],
    ];
}
