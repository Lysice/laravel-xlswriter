<?php

namespace Lysice\XlsWriter\Interfaces;

use Illuminate\Support\Collection;

interface FromCollection
{
    /**
     * @return Collection
     */
    public function collection();

    /**
     * @return array
     */
    public function headers() : array;
}
