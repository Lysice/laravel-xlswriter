<?php
namespace Lysice\XlsWriter\Facade;

use Illuminate\Support\Facades\Facade;

class Writer extends Facade
{
    protected static function getFacadeAccessor()
    {
        return \Lysice\XlsWriter\Writer::class;
    }
}
