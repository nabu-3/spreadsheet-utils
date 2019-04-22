<?php

/** @license
 *  Copyright 2019-2011 Rafael Gutierrez Martinez
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

namespace nabu\spreadsheet\exceptions;

use ErrorException;

use nabu\min\exceptions\ENabuException;

/**
 * Exception class to handle Spreadsheet exceptions.
 * @author Rafael Gutierrez <rgutierrez@nabu-3.com>
 * @since 0.0.2
 * @version 0.0.2
 * @package \nabu\lexer\exceptions
 */
class ENabuSpreadsheetUtilsException extends ENabuException
{
    /** @var int Invalid file name or path. Requires the file name and the IO error accessing or opening the file. */
    public const ERROR_INVALID_FILE_NAME_OR_PATH                    = 0x0001;
    /** @var int None Spreadsheet is loaded. */
    public const ERROR_NONE_SPREADSHEET_LOADED                      = 0x0002;

    /** @var array English error messages array. */
    private static $error_messages = array(
        ENabuSpreadsheetUtilsException::ERROR_INVALID_FILE_NAME_OR_PATH =>
            'Invalid name, file or path [%s].',
        ENabuSpreadsheetUtilsException::ERROR_NONE_SPREADSHEET_LOADED =>
            'None Spreadsheet is loaded.'
    );

    /**
     * Creates a Lexer Exception instance.
     * @param int $code Integer code of the exception.
     * @param array|null $values Valus to be inserted in the translated message if needed.
     * @throws ErrorException Trhos an exception if $code value is not supported.
     */
    public function __construct(int $code, array $values = null)
    {
        if (array_key_exists($code, self::$error_messages)) {
            parent::__construct(self::$error_messages[$code], $code, $values);
        } else {
            parent::__construct('Invalid exception code [%s]', 0, array($code));
        }
    }
}
