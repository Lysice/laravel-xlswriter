<?php

namespace Lysice\XlsWriter\Exceptions;

class TypeNotSupportException extends Exception
{
    public function __construct(
        $message = "The type passed doesn't not support. Make sure you either pass a valid extension to the filename or pass an explicit type.",
        $code = 0,
        \Throwable $previous = null
    ) {
        parent::__construct($message, $code, $previous);
    }
}
