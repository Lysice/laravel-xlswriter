<?php

namespace Lysice\XlsWriter\Interfaces;

use Lysice\XlsWriter\Supports\Format\DefaultFormat;
use Lysice\XlsWriter\Supports\Format\Format;

interface WithDefaultFormat {
    /**
     * @return array
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function defaultFormats() : array ;
}
