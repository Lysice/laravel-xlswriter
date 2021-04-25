<?php

namespace Lysice\XlsWriter\Interfaces;

interface FromArray
{
    /**
     * @return array
     */
    public function array(): array;

    /**
     * @return array
     */
    public function headers() : array;
}
