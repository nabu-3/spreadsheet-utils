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
 * @version 0.0.2
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
     * @param array $required_fields Array of field names that are mandatory before extract data.
     * @param bool $canonize If set to true, canonizes column names to remove unecessary lead paddings and special chars.
     * @return CNabuSpreadsheetData Returns an instance of @see { CNabuSpreadsheetData } with all extracted data.
     * @throws ENabuSpreadsheetUtilsException Throws an exception if something unexpected success.
     */
    public function extractColumns(
        array $translation_fields, array $required_fields, bool $canonize = false
    ): CNabuSpreadsheetData {
        if (is_null($this->spreadsheet)) {
            throw new ENabuSpreadsheetUtilsException(ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED);
        }

        $datasheet = $this->spreadsheet->getActiveSheet()->toArray(null, true, true, true);

        if (is_array($datasheet) && count($datasheet) > 0) {
            $translated_fields = $this->calculateColumnNameTranslations($translation_fields, $datasheet[1], $canonize);
            $this->checkMandatoryFields($translated_fields, $required_fields);
            $resultset = $this->mapData($datasheet, $translated_fields, $required_fields, 2);
        } else {
            $resultset = new CNabuSpreadsheetData();
        }

        return $resultset;
    }

    /**
     * Calculates the matrix to translate from initial column naming (A, B, C...) to final translation fields.
     * This method grants that columns could be unordered and extra columns interlaced between required columns.
     * @param array $translation_fields Associative array to translate fields from the first line column name
     * of the datasheet to the well formed name in the result data instance.
     * @param array $column_names Original column names found in the first line of the datasheet.
     * @param bool $canonize If set to true, canonizes column names to remove unecessary lead paddings and special chars.
     * @return array|null Returns an array with translation of columns.
     */
    private function calculateColumnNameTranslations(
        array $translation_fields, array $column_names, bool $canonize = false
    ): ?array {
        $keys = array_values($column_names);
        $values = array_keys($column_names);

        if ($canonize) {
            array_walk($keys, function(&$value) {
                $value = mb_strtolower(preg_replace('/_+$/', '', preg_replace('/^_+/', '', preg_replace('/[\s\.\(\)]+/', '_', $value))));
                $value = str_replace(SPSUTILS_VOCALS, SPSUTILS_CANONICAL, $value);
            });
        }

        $new_keys = array_intersect_key($translation_fields, array_combine($keys, $values));
        $new_values = array_intersect_key(array_combine($keys, $values), $translation_fields);

        $translated_columns = array();
        foreach ($new_keys as $key => $value) {
            $translated_columns[$value] = $new_values[$key];
        }

        return $translated_columns;
    }

    /**
     * Check if all mandatory fields are present in columns.
     * @param array $fields Fields found in columns.
     * @param array $required_fields List of mandatory field names.
     * @throws ENabuSpreadsheetUtilsException Throws an exception if some of mandatory fields are not present.
     */
    private function checkMandatoryFields(array $fields, array $required_fields): void
    {
        if (!(count($required_fields) === count(array_intersect(array_keys($fields), $required_fields)))) {
            $missed_columns = array_diff($required_fields, array_keys($fields));
            error_log(print_r($missed_columns, true));
            throw new ENabuSpreadsheetUtilsException(
                ENabuSpreadsheetUtilsException::ERROR_MANDATORY_COLUMNS_NOT_PRESENT,
                array(
                    implode(', ', $missed_columns)
                )
            );
        }
    }

    /**
     * Maps data in a @see { CNabuSpreadsheetData } instance and returns it.
     * @param array $datasheet Source datasheet.
     * @param array $map_fields Map columns correspondence to map data.
     * @param array $required_fields List of mandatory field names.
     * @param int $offset Offset of first line of data to map. Default value is 1 but the Datasheet array
     * of PhpSpreasheet is option base 1.
     * @return CNabuSpreadsheetData Returns the instance of @see { CNabuSpreadsheetData } with translated data.
     * @throws ENabuSpreadsheetUtilsException Throws an exception if something fails.
     */
    private function mapData(array $datasheet, array $map_fields, array $required_fields, int $offset = 1): CNabuSpreadsheetData
    {
        $resultset = new CNabuSpreadsheetData();

        if (($l = count($datasheet)) >= $offset) {
            for ($i = $offset; $i <= $l; $i++) {
                if (array_key_exists($i, $datasheet)) {
                    $source = $datasheet[$i];
                    $reg = array();
                    foreach ($map_fields as $key => $pos) {
                        $reg[$key] = $source[$pos];
                    }
                    if (count($missed_fields = array_diff(array_values($required_fields), array_keys($reg))) > 0) {
                        throw new ENabuSpreadsheetUtilsException(
                            ENabuSpreadsheetUtilsException::ERROR_MANDATORY_COLUMNS_NOT_PRESENT,
                            array(implode(', ', $missed_fields))
                        );
                    }
                    $resultset->setValue($i, $reg);
                }
            }
        }

        return $resultset;
    }
}
