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

use nabu\min\CNabuObject;

use nabu\spreadsheet\data\CNabuSpreadsheetData;

use nabu\spreadsheet\exceptions\ENabuSpreadsheetUtilsException;

use PhpOffice\PhpSpreadsheet;

/**
 * Class to read a Spreadsheet in MS(R) Office Excel format and convert to @see { TNabuSpreadsheetData } object.
 * @author Rafael Gutierrez <rgutierrez@nabu-3.com>
 * @since 0.0.1
 * @version 0.0.1
 * @package \nabu\spreadsheet
 */
class CNabuSpreadsheetReader extends CNabuObject
{
    /** @var string|null $filename Full path filename of MS(R) Excel file. */
    private $filename = null;
    /** @var PhpSpreadsheet|null $spreadsheet Native spreadsheet instance. */
    private $spreadsheet = null;

    /**
     * Creates the instance with the possibility to pass the file name of a MS(R) Excel file to open it.
     * @param string|null $filename Filename to read.
     */
    public function __construct(string $filename = null)
    {
        parent::__construct();

        if (is_string($filename)) {
            $this->loadFromFile($filename);
        }
    }

    /**
     * Validates a file name.
     * @param string $filename Filename to read.
     * @throws ENabuSpreadsheetUtilsException Throws an exception if something unexpected success opening the file.
     */
    private function validateFilename(string $filename): void
    {
        if (strlen($filename) === 0 ||
            ($this->filename = realpath($filename)) === false ||
            mb_strlen($this->filename) === 0 ||
            !file_exists($this->filename) ||
            !is_file($this->filename)
        ) {
            $this->filename = null;
            throw new ENabuSpreadsheetUtilsException(
                ENabuSpreadsheetUtilsException::ERROR_INVALID_FILE_NAME_OR_PATH,
                array($filename)
            );
        }
    }

    /**
     * Loads a MS(R) Excel file.
     * @param string $filename Filename to read.
     * @throws ENabuSpreadsheetUtilsException Throws an exception if something unexpected success opening the file.
     */
    public function loadFromFile(string $filename): void
    {
        $this->validateFilename($filename);
        $this->spreadsheet = IOFactory::load($this->filename);
        $this->setActiveSheetIndex(0);
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
     * @param bool $canonize If set to true, canonizes column names to remove unecessary lead paddings and special chars.
     * @return CNabuSpreadsheetData Returns an instance of @see { CNabuSpreadsheetData } with all extracted data.
     */
    public function extractColumns(array $translation_fields, bool $canonize = false): CNabuSpreadsheetData
    {
        if (is_null($this->spreadsheet)) {
            throw new ENabuSpreadsheetUtilsException(ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED);
        }

        $datasheet = $this->spreadsheet->getActiveSheet()->toArray(null, true, true, true);

        if (is_array($datasheet) && ($l = count($datasheet)) > 0) {
            $this->calculateColumnNameTranslations($translation_fields, $datasheet[1]);
        }

        return new CNabuSpreadsheetData();
    }

    /**
     * Calculates the matrix to translate from initial column naming (A, B, C...) to final translation fields.
     * This method grants that columns could be unordered and extra columns interlaced between required columns.
     * @param array $translation_fields Associative array to translate fields from the first line column name
     * of the datasheet to the well formed name in the result data instance.
     * @param array $column_names Original column names found in the first line of the datasheet.
     */
    private function calculateColumnNameTranslations(array $translation_fields, array $column_names)
    {
        $keys = array_values($column_names);
        $values = array_keys($column_names);

        array_walk($keys, function(&$value) {
            $value = mb_strtolower(preg_replace('/_+$/', '', preg_replace('/^_+/', '', preg_replace('/[\s\.\(\)]+/', '_', $value))));
            $value = str_replace(SPSUTILS_VOCALS, SPSUTILS_CANONICAL, $value);
        });

        $new_keys = array_intersect_key($translation_fields, array_combine($keys, $values));
        $new_values = array_intersect_key(array_combine($keys, $values), $translation_fields);
        $this->columns = array();
        foreach ($new_keys as $key => $value) {
            $this->columns[$value] = $new_values[$key];
        }

    }
}
