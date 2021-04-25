<?php

namespace Lysice\XlsWriter\Exceptions;

class ExportMustImplementException extends Exception
{
    public function __construct(
        $message = "Export class must implement interface: 'FromArray' or 'FromCollection'.",
        $code = 0,
        \Throwable $previous = null
    ) {
        parent::__construct($message, $code, $previous);
    }
}
