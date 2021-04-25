<?php

namespace Lysice\XlsWriter\Exceptions;

use InvalidArgumentException;
use Throwable;

class ParamErrorException extends InvalidArgumentException
{
    /**
     * @param string $message
     * @param int $code
     * @param Throwable|null $previous
     */
    public function __construct(
        $message = '',
        $code = 0,
        Throwable $previous = null
    ) {
        parent::__construct($message, $code, $previous);
    }

    /**
     * @return ParamErrorException
     */
    public static function import()
    {
        return new static('A filepath or UploadedFile needs to be passed to start the import.');
    }

    /**
     * @return ParamErrorException
     */
    public static function export()
    {
        return new static('A filepath needs to be passed in order to store the export.');
    }
}
