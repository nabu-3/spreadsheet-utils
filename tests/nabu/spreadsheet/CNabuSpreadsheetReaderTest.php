<?php

/** @license
 *  Copyright 2009-2011 Rafael Gutierrez Martinez
 *  Copyright 2012-2013 Welma WEB MKT LABS, S.L.
 *  Copyright 2014-2016 Where Ideas Simply Come True, S.L.
 *  Copyright 2017 nabu-3 Group
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */

namespace nabu\spreadsheet;

use PHPUnit\Framework\TestCase;

use nabu\spreadsheet\exceptions\ENabuSpreadsheetUtilsException;

/**
 * Tests for class @see { TNabuSpreadsheetData }.
 * @author Rafael Gutierrez <rgutierrez@nabu-3.com>
 * @since 0.0.1
 * @version 0.0.1
 * @package \nabu\spreadsheet
 */
class CNabuSpreadsheetReaderTest extends TestCase
{
    /**
     * @test __construct
     */
    public function testConstruct()
    {
        $reader = new CNabuSpreadsheetReader();
        $this->assertInstanceOf(CNabuSpreadsheetReader::class, $reader);

        $reader = new CNabuSpreadsheetReader(__DIR__ . DIRECTORY_SEPARATOR . 'resources/basic-excel-file.xlsx');
        $this->assertInstanceOf(CNabuSpreadsheetReader::class, $reader);

        $this->expectException(ENabuSpreadsheetUtilsException::class);
        $this->expectExceptionCode(ENabuSpreadsheetUtilsException::ERROR_INVALID_FILE_NAME_OR_PATH);
        $this->expectExceptionMessage('resources/not-exists.xlsx');
        $reader = new CNabuSpreadsheetReader(__DIR__ . DIRECTORY_SEPARATOR . 'resources/not-exists.xlsx');
    }

    /**
     * @test extractColumns
     */
    public function testExtractColumns()
    {
        $reader = new CNabuSpreadsheetReader(__DIR__ . DIRECTORY_SEPARATOR . 'resources/basic-excel-file.xlsx');
        $this->assertInstanceOf(CNabuSpreadsheetReader::class, $reader);

        $data = $reader->extractColumns(array(
            'column 1' => 'column_1',
            'column 2' => 'column_2',
            'column 3' => 'column_3',
            true
        ));
    }
}
