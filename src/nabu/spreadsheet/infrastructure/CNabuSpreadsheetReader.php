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

use PhpOffice\PhpSpreadsheet\IOFactory;

use nabu\data\interfaces\INabuDataList;

use nabu\infrastructure\reader\CNabuAbstractDataListFileReader;

use nabu\spreadsheet\data\CNabuSpreadsheetData;

use nabu\spreadsheet\exceptions\ENabuSpreadsheetUtilsException;

use PhpOffice\PhpSpreadsheet;

/**
 * Class to read a Spreadsheet in MS(R) Office Excel format and convert to @see { TNabuSpreadsheetData } object.
 * @author Rafael Gutierrez <rgutierrez@nabu-3.com>
 * @since 0.0.1
 * @version 0.0.2
 * @package \nabu\spreadsheet\infrastructure
 */
class CNabuSpreadsheetReader extends CNabuAbstractDataListFileReader
{
    /** @var string|null Index field name. */
    protected $index_field = null;

    /** @var PhpSpreadsheet|null $spreadsheet Native spreadsheet instance. */
    private $spreadsheet = null;

    protected function getValidMIMETypes(): array
    {
        return [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ];
    }

    protected function customFileValidation(string $filename): bool
    {
        return true;
    }

    protected function openSourceFile(string $filename): bool
    {
        $this->spreadsheet = IOFactory::load($filename);
        $this->setActiveSheetIndex(0);

        return true;
    }

    protected function closeSourceFile(): void
    {

    }

    protected function createDataListInstance(): INabuDataList
    {
        return new CNabuSpreadsheetData($this->index_field);
    }

    protected function getSourceDataAsArray(): ?array
    {
        return $this->spreadsheet->getActiveSheet()->toArray(null, true, true, true);
    }

    protected function checkBeforeParse(): bool
    {
        if (is_null($this->spreadsheet)) {
            throw new ENabuSpreadsheetUtilsException(ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED);
        }

        return true;
    }

    protected function checkAfterParse(\nabu\data\interfaces\INabuDataList $resultset): bool
    {
        return true;
    }

    /**
     * Get the Index Field used to index data in the result collection.
     * @return string|null Returns the index name if set or null otherwise.
     */
    public function getIndexField(): ?string
    {
        return $this->index_field;
    }

    /**
     * Set the Index Field used to index data in the result collection.
     * @param string|null $index The index field name in the result to use.
     * @return CNabuSpreadsheetReader Returns the self pointer to grant Fluent Interface.
     */
    public function setIndexField(?string $index = null): CNabuSpreadsheetReader
    {
        $this->index_field = $index;

        return $this;
    }

    /**
     * Set current Datasheet in the document.
     * @param int $index The index of datasheet to activate.
     * @return CNabuSpreadsheetReader Returns the self pointer to grant fluent interface.
     * @throws ENabuSpreadsheetUtilsException Throws an exception if something unexpected happens.
     */
    public function setActiveSheetIndex(int $index): CNabuSpreadsheetReader
    {
        if (is_null($this->spreadsheet)) {
            throw new ENabuSpreadsheetUtilsException(ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED);
        }

        $this->spreadsheet->setActiveSheetIndex($index);

        return $this;
    }


}
