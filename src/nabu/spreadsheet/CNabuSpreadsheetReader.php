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

use PhpOffice\PhpSpreadsheet\IOFactory;

use nabu\infrastructure\reader\CNabuAbstractDataListFileReader;

use nabu\spreadsheet\data\CNabuSpreadsheetData;
use nabu\spreadsheet\data\CNabuSpreadsheetDataRecord;

use nabu\spreadsheet\exceptions\ENabuSpreadsheetUtilsException;

use PhpOffice\PhpSpreadsheet;

/**
 * Class to read a Spreadsheet in MS(R) Office Excel format and convert to @see { TNabuSpreadsheetData } object.
 * @author Rafael Gutierrez <rgutierrez@nabu-3.com>
 * @since 0.0.1
 * @version 0.0.2
 * @package \nabu\spreadsheet
 */
class CNabuSpreadsheetReader extends CNabuAbstractDataListFileReader
{
    /** @var PhpSpreadsheet|null $spreadsheet Native spreadsheet instance. */
    private $spreadsheet = null;

    protected function getValidMIMETypes(): array
    {
        throw new \LogicException('Not implemented'); // TODO
    }

    protected function customFileValidation(string $filename): bool
    {
        return true;
    }

    protected function openSourceFile(string $filename): bool
    {
        $this->spreadsheet = IOFactory::load($this->filename);
        $this->setActiveSheetIndex(0);

        return true;
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

    /**
     * Read data as a massive process that generates a @see { CNabuSpreadsheetData } instance with available data.
     * This process discards not required columns in the datasheet and renames first line to well formed field names.
     * @param array $translation_fields Associative array to translate fields from the first line column name
     * of the datasheet to the well formed name in the result data instance.
     * @param array $required_fields Array of field names that are mandatory before extract data.
     * @param string|null $index_field If setted, the list is indexed using values in this field.
     * @param bool $canonize If set to true, canonizes column names to remove unecessary lead paddings and special chars.
     * @return CNabuSpreadsheetData Returns an instance of @see { CNabuSpreadsheetData } with all extracted data.
     * @throws ENabuSpreadsheetUtilsException Throws an exception if something unexpected success.
     */
    public function extractColumns(
        array $translation_fields, array $required_fields, ?string $index_field = null, bool $canonize = false
    ): CNabuSpreadsheetData {
        if (is_null($this->spreadsheet)) {
            throw new ENabuSpreadsheetUtilsException(ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED);
        }

        $datasheet = $this->spreadsheet->getActiveSheet()->toArray(null, true, false, true);

        if (is_array($datasheet) && count($datasheet) > 0) {
            if (is_string($index_field) && !in_array($index_field, $required_fields)) {
                $required_fields[] = $index_field;
            }
            $translated_fields = $this->calculateColumnNameTranslations($translation_fields, $datasheet[1], $canonize);
            $this->checkMandatoryFields($translated_fields, $required_fields);
            $resultset = $this->mapData($datasheet, $translated_fields, $required_fields, $index_field, 2);
        } else {
            $resultset = $this->createDataInstance();
        }

        return $resultset;
    }

    /**
     * This method is called internally to create the CNabuSpreadsheetData instance. You can override it to return
     * a subclass of @see { CNabuSpreasheetData } class.
     * @param string|null $index_field Index field to index data collection.
     * @return CNabuSpreadsheetData Returns the new instance.
     */
    protected function createDataInstance(?string $index_field = null): CNabuSpreadsheetData
    {
        return new CNabuSpreadsheetData($index_field);
    }


}
