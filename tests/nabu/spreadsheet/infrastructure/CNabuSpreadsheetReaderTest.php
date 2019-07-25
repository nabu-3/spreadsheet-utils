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

namespace nabu\spreadsheet\infrastructure;

use PHPUnit\Framework\Error\Error;

use PHPUnit\Framework\TestCase;

use nabu\infrastructure\reader\interfaces\INabuDataListReader;
use nabu\infrastructure\reader\interfaces\INabuDataListFileReader;

use nabu\spreadsheet\exceptions\ENabuSpreadsheetUtilsException;

/**
 * Tests for class @see { TNabuSpreadsheetData }.
 * @author Rafael Gutierrez <rgutierrez@nabu-3.com>
 * @since 0.0.1
 * @version 0.0.2
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
        $this->assertInstanceOf(INabuDataListReader::class, $reader);
        $this->assertInstanceOf(INabuDataListFileReader::class, $reader);

        $reader = new CNabuSpreadsheetReader(__DIR__ . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'basic-excel-file.xlsx');
        $this->assertInstanceOf(CNabuSpreadsheetReader::class, $reader);
        $this->assertInstanceOf(INabuDataListReader::class, $reader);
        $this->assertInstanceOf(INabuDataListFileReader::class, $reader);

        $this->expectException(Error::class);
        $this->expectExceptionMessage('resources/not-exists.xlsx');
        $reader = new CNabuSpreadsheetReader(__DIR__ . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'not-exists.xlsx');
    }

    /**
     * @test parse
     */
    public function testParseOrdinalArray()
    {
        $reader = new CNabuSpreadsheetReader(
            __DIR__ . DIRECTORY_SEPARATOR . 'resources/basic-excel-file.xlsx',
            array(
                'column_2' => 'value_1',
                'column_3' => 'value_2',
                'column_1' => 'value_3'
            ),
            array(
                'value_1', 'value_2'
            ),
            false, 1, 2
        );
        $this->assertInstanceOf(CNabuSpreadsheetReader::class, $reader);

        $data = $reader->parse();

        $this->assertTrue($data->getItem(0)->isValueEqualTo('value_1', 'Test string'));
        $this->assertTrue($data->getItem(0)->isValueEqualTo('value_2', 369));
        $this->assertTrue($data->getItem(0)->isValueEqualTo('value_3', 123));

        $this->assertTrue($data->getItem(1)->isValueEqualTo('value_1', 'Other test'));
        $this->assertTrue($data->getItem(1)->isValueEqualTo('value_2', 258));
        $this->assertTrue($data->getItem(1)->isValueEqualTo('value_3', 456));

        $this->assertTrue($data->getItem(2)->isValueEqualTo('value_1', 'More tests data'));
        $this->assertTrue($data->getItem(2)->isValueEqualTo('value_2', 147));
        $this->assertTrue($data->getItem(2)->isValueEqualTo('value_3', 789));
    }

    /**
     * @test extractColumns
     */
    public function testParseIndexedArray()
    {
        $reader = new CNabuSpreadsheetReader(__DIR__ . DIRECTORY_SEPARATOR . 'resources/basic-excel-file.xlsx');
        $this->assertInstanceOf(CNabuSpreadsheetReader::class, $reader);

        $reader->setConvertFieldsMatrix(array(
            'column_2' => 'value_1',
            'column_3' => 'value_2',
            'column_1' => 'value_3'
        ));

        $reader->setRequiredFields(array(
            'value_1', 'value_2'
        ));

        $reader->setUseStrictSourceNames(false);
        $reader->setHeaderNamesOffset(1);
        $reader->setFirstRowOffset(2);
        $reader->setIndexField('value_2');
        $this->assertSame('value_2', $reader->getIndexField());

        $data = $reader->parse();

        $this->assertTrue($data->getItem(369)->isValueEqualTo('value_1', 'Test string'));
        $this->assertTrue($data->getItem(369)->isValueEqualTo('value_2', 369));
        $this->assertTrue($data->getItem(369)->isValueEqualTo('value_3', 123));
        $this->assertSame(
            array(
                'value_1' => 'Test string',
                'value_2' => 369.0,
                'value_3' => 123.0
            ),
            $data->getItem(369)->getValuesAsArray()
        );

        $this->assertTrue($data->getItem(258)->isValueEqualTo('value_1', 'Other test'));
        $this->assertTrue($data->getItem(258)->isValueEqualTo('value_2', 258));
        $this->assertTrue($data->getItem(258)->isValueEqualTo('value_3', 456));
        $this->assertSame(
            array(
                'value_1' => 'Other test',
                'value_2' => 258.0,
                'value_3' => 456.0
            ),
            $data->getItem(258)->getValuesAsArray()
        );

        $this->assertTrue($data->getItem(147)->isValueEqualTo('value_1', 'More tests data'));
        $this->assertTrue($data->getItem(147)->isValueEqualTo('value_2', 147));
        $this->assertTrue($data->getItem(147)->isValueEqualTo('value_3', 789));
        $this->assertSame(
            array(
                'value_1' => 'More tests data',
                'value_2' => 147.0,
                'value_3' => 789.0
            ),
            $data->getItem(147)->getValuesAsArray()
        );

        $this->assertNull($data->getItem('another'));
    }

    /**
     * @test checkBeforeParse
     */
    public function testCheckBeforeParseFails()
    {
        $reader = new CNabuSpreadsheetReader();

        $this->expectException(ENabuSpreadsheetUtilsException::class);
        $this->expectExceptionCode(ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED);

        $reader->parse();
    }

    /**
     * @test setActiveSheetIndex
     */
    public function testSetActiveSheetIndexFails()
    {
        $reader = new CNabuSpreadsheetReader();

        $this->expectException(ENabuSpreadsheetUtilsException::class);
        $this->expectExceptionCode(ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED);

        $reader->setActiveSheetIndex(1);
    }
}
