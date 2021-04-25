<?php

namespace Lysice\XlsWriter\Interfaces;
use Illuminate\Database\Query\Builder;

interface FromQuery
{
    /**
     * @return Builder
     */
    public function query();

    public function headers() : array ;
}
